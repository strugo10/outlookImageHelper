# Skrypt weryfikujacy poprawnosc skladniowa oraz analiza AST (F821 / F811) wszystkich plikow Pythona (.py)
# przy uzyciu rzeczywistego silnika Python z NVDA (C:\Program Files\NVDA\python313.dll).
# Zwraca kod wyjscia 0 jesli kod jest w 100% poprawny, lub 1 jesli wystapil blad.

param(
    [Parameter(Mandatory=$false)]
    [string]$TargetDir = (Get-Location).Path
)

$ErrorActionPreference = "Stop"

$pyFiles = Get-ChildItem -Path $TargetDir -Recurse -Filter "*.py"
if (-not $pyFiles) {
    Write-Host "Brak plikow Python (.py) do weryfikacji w: $TargetDir" -ForegroundColor Yellow
    exit 0
}

Write-Host "Weryfikacja skladni Python i analiza AST dla $($pyFiles.Count) plikow w $TargetDir..." -ForegroundColor Cyan

$fileListLiteral = ($pyFiles | ForEach-Object { "    r'''$($_.FullName)'''," }) -join "`r`n"

$tempRunner = Join-Path $env:TEMP "nvda_syntax_validator_$PID.py"
$runnerTemplate = @'
import sys, os, py_compile, ast, builtins

files = [
__FILES_PLACEHOLDER__
]

BUILTIN_NAMES = set(dir(builtins)) | {"_", "gettext", "ngettext", "pgettext", "npgettext", "__file__", "__name__", "__doc__", "__package__", "__spec__", "__annotations__"}

def check_ast_diagnostics(filepath):
    norm_path = filepath.replace("/", "\\").lower()
    if "\\lib\\" in norm_path:
        return []

    with open(filepath, "r", encoding="utf-8-sig") as fh:
        source = fh.read()
    tree = ast.parse(source, filename=filepath)
    errors = []

    has_future_annotations = False
    for node in tree.body:
        if isinstance(node, ast.ImportFrom) and node.module == "__future__":
            if any(alias.name == "annotations" for alias in node.names):
                has_future_annotations = True

    # 1. Check top-level duplicate function/class definitions and duplicate imports (F811)
    top_defs = {}
    top_imports = {}
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            if node.name in top_defs:
                errors.append(f"Linia {node.lineno}: F811 Zduplikowana definicja '{node.name}' (poprzednia w linii {top_defs[node.name]})")
            top_defs[node.name] = node.lineno
        elif isinstance(node, ast.Import):
            for alias in node.names:
                bname = alias.asname or alias.name.split(".")[0]
                if bname in top_imports and top_imports[bname][1] == (alias.name, alias.asname):
                    errors.append(f"Linia {node.lineno}: F811 Zduplikowany import '{bname}' (poprzedni w linii {top_imports[bname][0]})")
                top_imports[bname] = (node.lineno, (alias.name, alias.asname))
        elif isinstance(node, ast.ImportFrom):
            for alias in node.names:
                if alias.name == "*":
                    continue
                bname = alias.asname or alias.name
                key = (node.module, alias.name, alias.asname)
                if bname in top_imports and top_imports[bname][1] == key:
                    errors.append(f"Linia {node.lineno}: F811 Zduplikowany import '{bname}' (poprzedni w linii {top_imports[bname][0]})")
                top_imports[bname] = (node.lineno, key)

    # 2. Scope analysis for undefined names (F821)
    def collect_scope_defs(stmts, args_obj=None):
        defs = set()
        globals_decl = set()
        nonlocals_decl = set()
        if args_obj:
            for a in (args_obj.posonlyargs + args_obj.args + args_obj.kwonlyargs):
                defs.add(a.arg)
            if args_obj.vararg:
                defs.add(args_obj.vararg.arg)
            if args_obj.kwarg:
                defs.add(args_obj.kwarg.arg)

        def scan_target(t):
            if isinstance(t, ast.Name):
                defs.add(t.id)
            elif isinstance(t, (ast.Tuple, ast.List)):
                for elt in t.elts:
                    scan_target(elt)
            elif isinstance(t, ast.Starred):
                scan_target(t.value)

        def walk_stmt(s):
            if isinstance(s, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
                defs.add(s.name)
            elif isinstance(s, (ast.Import, ast.ImportFrom)):
                for alias in s.names:
                    if alias.name == "*":
                        defs.add("*")
                    else:
                        defs.add(alias.asname or alias.name.split(".")[0])
            elif isinstance(s, ast.Assign):
                for t in s.targets:
                    scan_target(t)
                walk_expr_for_named(s.value)
            elif isinstance(s, ast.AnnAssign):
                if s.target:
                    scan_target(s.target)
                if s.value:
                    walk_expr_for_named(s.value)
            elif isinstance(s, ast.AugAssign):
                scan_target(s.target)
                walk_expr_for_named(s.value)
            elif isinstance(s, (ast.For, ast.AsyncFor)):
                scan_target(s.target)
                for sub in s.body + s.orelse:
                    walk_stmt(sub)
            elif isinstance(s, (ast.With, ast.AsyncWith)):
                for item in s.items:
                    if item.optional_vars:
                        scan_target(item.optional_vars)
                for sub in s.body:
                    walk_stmt(sub)
            elif isinstance(s, ast.Try):
                for sub in s.body + s.orelse + s.finalbody:
                    walk_stmt(sub)
                for h in s.handlers:
                    if h.name:
                        defs.add(h.name)
                    for sub in h.body:
                        walk_stmt(sub)
            elif isinstance(s, ast.If):
                walk_expr_for_named(s.test)
                for sub in s.body + s.orelse:
                    walk_stmt(sub)
            elif isinstance(s, ast.While):
                walk_expr_for_named(s.test)
                for sub in s.body + s.orelse:
                    walk_stmt(sub)
            elif isinstance(s, ast.Global):
                globals_decl.update(s.names)
            elif isinstance(s, ast.Nonlocal):
                nonlocals_decl.update(s.names)
            elif isinstance(s, ast.Delete):
                for t in s.targets:
                    if isinstance(t, ast.Name):
                        defs.add(t.id)
            else:
                walk_expr_for_named(s)

        def walk_expr_for_named(node):
            for child in ast.iter_child_nodes(node):
                if isinstance(child, (ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef, ast.Lambda, ast.ListComp, ast.SetComp, ast.DictComp, ast.GeneratorExp)):
                    continue
                if isinstance(child, ast.NamedExpr) and isinstance(child.target, ast.Name):
                    defs.add(child.target.id)
                walk_expr_for_named(child)

        for st in stmts:
            walk_stmt(st)
        return defs, globals_decl, nonlocals_decl

    module_defs, _, _ = collect_scope_defs(tree.body)

    def check_node(node, visible_scopes, is_class_scope=False):
        if isinstance(node, ast.Name) and isinstance(node.ctx, ast.Load):
            name = node.id
            if name not in BUILTIN_NAMES and name not in module_defs:
                if not any((name in sc or "*" in sc) for sc, _ in visible_scopes):
                    if "*" not in module_defs:
                        errors.append(f"Linia {node.lineno}: F821 Niezdefiniowana nazwa '{name}'")
            return

        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            for dec in node.decorator_list:
                check_node(dec, visible_scopes, is_class_scope)
            for d in node.args.defaults + [x for x in node.args.kw_defaults if x is not None]:
                check_node(d, visible_scopes, is_class_scope)
            if not has_future_annotations:
                for a in (node.args.posonlyargs + node.args.args + node.args.kwonlyargs):
                    if a.annotation:
                        check_node(a.annotation, visible_scopes, is_class_scope)
                if node.args.vararg and node.args.vararg.annotation:
                    check_node(node.args.vararg.annotation, visible_scopes, is_class_scope)
                if node.args.kwarg and node.args.kwarg.annotation:
                    check_node(node.args.kwarg.annotation, visible_scopes, is_class_scope)
                if node.returns:
                    check_node(node.returns, visible_scopes, is_class_scope)
            f_defs, f_globals, _ = collect_scope_defs(node.body, node.args)
            module_defs.update(f_globals)
            new_scopes = [(sc, False) for sc, is_cls in visible_scopes if not is_cls] + [(f_defs, False)]
            for st in node.body:
                check_node(st, new_scopes, False)
            return

        if isinstance(node, ast.AnnAssign):
            if node.target:
                check_node(node.target, visible_scopes, is_class_scope)
            if not has_future_annotations and node.annotation:
                check_node(node.annotation, visible_scopes, is_class_scope)
            if node.value:
                check_node(node.value, visible_scopes, is_class_scope)
            return

        if isinstance(node, ast.ClassDef):
            for dec in node.decorator_list:
                check_node(dec, visible_scopes, is_class_scope)
            for b in node.bases:
                check_node(b, visible_scopes, is_class_scope)
            for kw in node.keywords:
                check_node(kw.value, visible_scopes, is_class_scope)
            c_defs, c_globals, _ = collect_scope_defs(node.body)
            module_defs.update(c_globals)
            new_scopes = visible_scopes + [(c_defs, True)]
            for st in node.body:
                check_node(st, new_scopes, True)
            return

        if isinstance(node, ast.Lambda):
            for d in node.args.defaults + [x for x in node.args.kw_defaults if x is not None]:
                check_node(d, visible_scopes, is_class_scope)
            l_defs, _, _ = collect_scope_defs([], node.args)
            new_scopes = [(sc, False) for sc, is_cls in visible_scopes if not is_cls] + [(l_defs, False)]
            check_node(node.body, new_scopes, False)
            return

        if isinstance(node, (ast.ListComp, ast.SetComp, ast.GeneratorExp, ast.DictComp)):
            comp_defs = set()
            def scan_comp_target(t):
                if isinstance(t, ast.Name):
                    comp_defs.add(t.id)
                elif isinstance(t, (ast.Tuple, ast.List)):
                    for e in t.elts:
                        scan_comp_target(e)
                elif isinstance(t, ast.Starred):
                    scan_comp_target(t.value)
            cur_scopes = [(sc, False) for sc, is_cls in visible_scopes if not is_cls]
            for idx, gen in enumerate(node.generators):
                if idx == 0:
                    check_node(gen.iter, visible_scopes, is_class_scope)
                else:
                    check_node(gen.iter, cur_scopes + [(comp_defs, False)], False)
                scan_comp_target(gen.target)
                for if_clause in gen.ifs:
                    check_node(if_clause, cur_scopes + [(comp_defs, False)], False)
            if isinstance(node, ast.DictComp):
                check_node(node.key, cur_scopes + [(comp_defs, False)], False)
                check_node(node.value, cur_scopes + [(comp_defs, False)], False)
            else:
                check_node(node.elt, cur_scopes + [(comp_defs, False)], False)
            return

        for child in ast.iter_child_nodes(node):
            check_node(child, visible_scopes, is_class_scope)

    for st in tree.body:
        check_node(st, [(module_defs, False)], False)

    return errors

failed = False
for f in files:
    try:
        py_compile.compile(f, doraise=True)
        ast_errs = check_ast_diagnostics(f)
        if ast_errs:
            for err in ast_errs:
                print(f"  [BLAD AST] {f}: {err}", file=sys.stderr)
            failed = True
        else:
            print(f"  [OK] {f}")
    except Exception as e:
        print(f"  [BLAD SKLADNI] {f}: {e}", file=sys.stderr)
        failed = True

if failed:
    sys.exit(1)
print("Wszystkie pliki Python przeszly walidacje skladni oraz AST pomyslnie!")
sys.exit(0)
'@

$runnerCode = $runnerTemplate.Replace("__FILES_PLACEHOLDER__", $fileListLiteral)
[System.IO.File]::WriteAllText($tempRunner, $runnerCode, [System.Text.Encoding]::UTF8)

$env:PYTHONHOME = "C:\Program Files\NVDA"
$env:PYTHONPATH = "C:\Program Files\NVDA\library.zip;C:\Program Files\NVDA"

try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public class NvdaSyntaxValidator {
    [DllImport(@"C:\Program Files\NVDA\python313.dll", EntryPoint = "Py_Initialize")]
    public static extern void Py_Initialize();

    [DllImport(@"C:\Program Files\NVDA\python313.dll", EntryPoint = "PyRun_SimpleString")]
    public static extern int PyRun_SimpleString(string code);
}
"@ -ErrorAction SilentlyContinue
} catch {}

[NvdaSyntaxValidator]::Py_Initialize()
$execCmd = "import runpy, sys; runpy.run_path(r'''$tempRunner''')"
$res = [NvdaSyntaxValidator]::PyRun_SimpleString($execCmd)

Remove-Item $tempRunner -Force -ErrorAction SilentlyContinue

if ($res -ne 0) {
    Write-Host "Walidacja Python zakonczyla sie BLEDEM (kod $res)! Budowanie zablokowane." -ForegroundColor Red
    exit 1
} else {
    Write-Host "Walidacja skladni i AST zakonczona sukcesem." -ForegroundColor Green
    exit 0
}
