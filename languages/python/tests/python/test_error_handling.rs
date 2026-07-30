use super::helpers::*;

// ══════════════════════════════════════════════════════════════════════════════
// Exception handling extended
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn except_tuple_of_types() {
    compile_ok("try:\n    pass\nexcept (ValueError, TypeError):\n    pass\n");
}

#[test]
fn except_tuple_with_as() {
    compile_ok("try:\n    pass\nexcept (ValueError, TypeError) as e:\n    print(e)\n");
}

#[test]
fn try_finally_without_except() {
    compile_ok("try:\n    f = open('x')\nfinally:\n    f.close()\n");
}

#[test]
fn bare_raise_reraise() {
    compile_ok("try:\n    x = 1 / 0\nexcept:\n    print('logging')\n    raise\n");
}

#[test]
fn except_hierarchy() {
    compile_ok(
        "try:\n    pass\nexcept ValueError:\n    pass\nexcept Exception:\n    pass\nexcept:\n    pass\n",
    );
}

#[test]
fn raise_from() {
    compile_ok(
        "try:\n    pass\nexcept Exception as e:\n    raise RuntimeError('wrapped') from e\n",
    );
}

#[test]
fn nested_try_except() {
    compile_ok(
        "try:\n    try:\n        risky()\n    except ValueError:\n        pass\nexcept Exception:\n    pass\n",
    );
}

// ══════════════════════════════════════════════════════════════════════════════
// Control flow extended
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn for_else_no_break() {
    let out = run_python(
        "for x in [1, 2, 3]:\n    if x == 5:\n        break\nelse:\n    print('no five')\n",
    );
    assert_eq!(out[0], "no five");
}

#[test]
fn for_else_with_break() {
    let out = run_python(
        "for x in [1, 2, 3]:\n    if x == 2:\n        break\nelse:\n    print('no break')\nprint('done')\n",
    );
    assert_eq!(out[0], "done");
}

#[test]
fn while_else_no_break() {
    let out = run_python("i = 0\nwhile i < 3:\n    i += 1\nelse:\n    print('completed')\n");
    assert_eq!(out[0], "completed");
}

#[test]
fn nested_loops_break_continue() {
    let out = run_python(
        r#"
for i in range(3):
    for j in range(3):
        if j == 1:
            continue
        if i == 2:
            break
        print(i, j)
"#,
    );
    assert_eq!(out, vec!["0 0", "0 2", "1 0", "1 2"]);
}

// ══════════════════════════════════════════════════════════════════════════════
// Builtins runtime
// ══════════════════════════════════════════════════════════════════════════════

#[test]
fn builtin_len() {
    assert_eq!(run_python_one("print(len([1, 2, 3]))\n"), "3");
}

#[test]
fn builtin_abs() {
    assert_eq!(run_python_one("print(abs(-5))\n"), "5");
}

#[test]
fn builtin_int_conversion() {
    assert_eq!(run_python_one("print(int('42'))\n"), "42");
}

#[test]
fn builtin_str_conversion() {
    assert_eq!(run_python_one("print(str(42))\n"), "42");
}

#[test]
fn builtin_bool_falsy() {
    assert_eq!(run_python_one("print(bool(0))\n"), "False");
}

#[test]
fn builtin_bool_truthy() {
    assert_eq!(run_python_one("print(bool(1))\n"), "True");
}

#[test]
fn builtin_sum() {
    assert_eq!(run_python_one("print(sum([1, 2, 3, 4, 5]))\n"), "15");
}

#[test]
fn builtin_min() {
    assert_eq!(run_python_one("print(min(3, 1, 2))\n"), "1");
}

#[test]
fn builtin_max() {
    assert_eq!(run_python_one("print(max(3, 1, 2))\n"), "3");
}

#[test]
fn builtin_sorted_runtime() {
    compile_ok("x = sorted([3, 1, 2])\n");
}

#[test]
fn builtin_range_two_args() {
    let out = run_python("for i in range(2, 5):\n    print(i)\n");
    assert_eq!(out, vec!["2", "3", "4"]);
}

#[test]
fn builtin_range_three_args() {
    let out = run_python("for i in range(0, 10, 3):\n    print(i)\n");
    assert_eq!(out, vec!["0", "3", "6", "9"]);
}

#[test]
fn builtin_enumerate_runtime() {
    let out = run_python("for i, v in enumerate(['a', 'b']):\n    print(i, v)\n");
    assert_eq!(out, vec!["0 a", "1 b"]);
}

#[test]
fn builtin_zip_runtime() {
    let out = run_python("for a, b in zip([1, 2], [3, 4]):\n    print(a, b)\n");
    assert_eq!(out, vec!["1 3", "2 4"]);
}

#[test]
fn builtin_chr() {
    assert_eq!(run_python_one("print(chr(65))\n"), "A");
}

#[test]
fn builtin_ord() {
    assert_eq!(run_python_one("print(ord('A'))\n"), "65");
}

#[test]
fn builtin_type() {
    compile_ok("print(type(42))\n");
}

#[test]
fn builtin_isinstance() {
    compile_ok("print(isinstance(42, int))\n");
}

#[test]
fn builtin_reversed() {
    compile_ok("for x in reversed([1, 2, 3]):\n    print(x)\n");
}

#[test]
fn builtin_map_filter() {
    compile_ok("result = list(map(str, [1, 2, 3]))\nresult2 = list(filter(None, [0, 1, 2]))\n");
}
