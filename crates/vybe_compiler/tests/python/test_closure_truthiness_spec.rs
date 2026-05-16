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

runtime_case!(closure_capture_runtime, "def outer(x):\n    def inner(y):\n        return x + y\n    return inner\nfn = outer(10)\nprint(fn(5))\n", "15");
runtime_case!(closure_mutable_state_runtime, "def outer():\n    total = [0]\n    def inner():\n        total[0] += 1\n        return total[0]\n    return inner\nfn = outer()\nprint(fn())\n", "1");
runtime_case!(nonlocal_increment_runtime, "def outer():\n    x = 0\n    def inner():\n        nonlocal x\n        x += 1\n        return x\n    return inner\nfn = outer()\nprint(fn())\n", "1");
runtime_case!(nested_shadow_runtime, "x = 'global'\ndef outer():\n    x = 'outer'\n    def inner():\n        x = 'inner'\n        return x\n    return inner()\nprint(outer())\n", "inner");
runtime_case!(global_write_runtime, "x = 1\ndef bump():\n    global x\n    x = x + 1\nbump()\nprint(x)\n", "2");
runtime_case!(lambda_closure_runtime, "def outer(x):\n    return lambda y: x * y\nprint(outer(3)(4))\n", "12");
runtime_case!(closure_returns_function_runtime, "def outer():\n    msg = 'hi'\n    def inner():\n        return msg\n    return inner\nprint(outer()())\n", "hi");
compile_case!(closure_default_bind_compile, "funcs = [lambda x=i: x for i in range(3)]\n");
compile_case!(closure_loop_capture_compile, "funcs = []\nfor i in range(3):\n    funcs.append(lambda: i)\n");
runtime_case!(truth_empty_list_runtime, "if []:\n    print('yes')\nelse:\n    print('no')\n", "no");
runtime_case!(truth_nonempty_list_runtime, "if [1]:\n    print('yes')\nelse:\n    print('no')\n", "yes");
runtime_case!(truth_empty_dict_runtime, "if {}:\n    print('yes')\nelse:\n    print('no')\n", "no");
runtime_case!(truth_empty_string_runtime, "if '':\n    print('yes')\nelse:\n    print('no')\n", "no");
runtime_case!(truth_zero_runtime, "if 0:\n    print('yes')\nelse:\n    print('no')\n", "no");
runtime_case!(truth_one_runtime, "if 1:\n    print('yes')\nelse:\n    print('no')\n", "yes");
compile_case!(truth_none_in_if_compile, "if None:\n    pass\n");
compile_case!(bool_from_len_compile, "class Box:\n    def __len__(self):\n        return 1\n");
runtime_case!(and_short_circuit_runtime, "print(0 and 5)\n", "0");
runtime_case!(or_short_circuit_runtime, "print(0 or 5)\n", "5");
runtime_case!(ternary_truthiness_runtime, "x = []\nprint('a' if x else 'b')\n", "b");
runtime_case!(chained_or_runtime, "print('' or 0 or 'fallback')\n", "fallback");
runtime_case!(chained_and_runtime, "print(1 and 2 and 3)\n", "3");
runtime_case!(not_empty_list_runtime, "print(not [1])\n", "false");
compile_case!(bool_custom_dunder_compile, "class Flag:\n    def __bool__(self):\n        return False\n");
compile_case!(bool_custom_len_compile, "class Flag:\n    def __len__(self):\n        return 0\n");
compile_case!(nonlocal_two_levels_compile, "def outer():\n    x = 0\n    def mid():\n        def inner():\n            nonlocal x\n            x += 1\n        inner()\n");
compile_case!(closure_comprehension_compile, "funcs = [lambda y, x=x: x + y for x in range(3)]\n");
runtime_case!(short_circuit_side_effect_runtime, "calls = []\ndef mark():\n    calls.append(1)\n    return True\n0 and mark()\nprint(len(calls))\n", "0");
runtime_case!(is_none_runtime, "x = None\nprint(x is None)\n", "true");
runtime_case!(is_not_none_runtime, "x = 1\nprint(x is not None)\n", "true");