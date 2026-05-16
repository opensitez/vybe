use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

compile_case!(globals_builtin_compile, "x = 1\ng = globals()\n");
compile_case!(locals_builtin_compile, "def f():\n    x = 1\n    return locals()\n");
compile_case!(vars_builtin_compile, "class C:\n    pass\nc = C()\nv = vars(c)\n");
compile_case!(dir_builtin_compile, "names = dir(str)\n");
compile_case!(debug_dunder_compile, "flag = __debug__\n");
compile_case!(annotations_dunder_compile, "x: int = 1\nann = __annotations__\n");
compile_case!(module_dict_compile, "d = globals()['__name__']\n");
compile_case!(breakpoint_compile, "breakpoint()\n");
compile_case!(callable_builtin_compile, "def f():\n    pass\nok = callable(f)\n");
compile_case!(getattr_builtin_compile, "x = getattr(obj, 'name', None)\n");
compile_case!(setattr_builtin_compile, "setattr(obj, 'name', 1)\n");
compile_case!(delattr_builtin_compile, "delattr(obj, 'name')\n");
compile_case!(hasattr_builtin_compile, "ok = hasattr(obj, 'name')\n");
compile_case!(isinstance_tuple_compile, "ok = isinstance(1, (int, float))\n");
compile_case!(issubclass_tuple_compile, "ok = issubclass(bool, (int, float))\n");
runtime_case!(debug_dunder_runtime, "print(__debug__)\n", "true");
runtime_case!(name_dunder_runtime, "print(__name__)\n", "__main__");
runtime_case!(globals_lookup_runtime, "x = 42\nprint(globals()['x'])\n", "42");
runtime_case!(callable_runtime, "def f():\n    pass\nprint(callable(f))\n", "true");
compile_case!(repr_builtin_compile, "text = repr({'a': 1})\n");