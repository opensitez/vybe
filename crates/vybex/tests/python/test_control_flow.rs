use super::helpers::{run_python, compile_ok};

#[test]
fn if_simple() {
    compile_ok("if True:\n    print(1)\n");
}

#[test]
fn if_else() {
    compile_ok("if x:\n    print(1)\nelse:\n    print(2)\n");
}

#[test]
fn if_elif_else() {
    compile_ok("if a:\n    pass\nelif b:\n    pass\nelse:\n    pass\n");
}

#[test]
fn while_loop() {
    compile_ok("while True:\n    break\n");
}

#[test]
fn while_with_continue() {
    compile_ok("i = 0\nwhile i < 10:\n    i += 1\n    if i == 5:\n        continue\n    print(i)\n");
}

#[test]
fn for_range() {
    compile_ok("for i in range(10):\n    print(i)\n");
}

#[test]
fn for_list() {
    compile_ok("for x in [1, 2, 3]:\n    print(x)\n");
}

#[test]
fn for_with_break() {
    compile_ok("for i in range(10):\n    if i == 5:\n        break\n    print(i)\n");
}

#[test]
fn nested_loops() {
    compile_ok("for i in range(3):\n    for j in range(3):\n        print(i, j)\n");
}

#[test]
fn try_except() {
    compile_ok("try:\n    x = 1\nexcept:\n    print('error')\n");
}

#[test]
fn try_except_finally() {
    compile_ok("try:\n    x = 1\nexcept:\n    pass\nfinally:\n    print('done')\n");
}

#[test]
fn raise_exception() {
    compile_ok("raise ValueError()\n");
}

// Runtime control flow tests

#[test]
fn if_else_runtime() {
    let out = run_python("x = 10\nif x > 5:\n    print(\"big\")\nelse:\n    print(\"small\")\n");
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_elif_else_runtime() {
    let out = run_python("x = 5\nif x > 10:\n    print(\"big\")\nelif x > 3:\n    print(\"medium\")\nelse:\n    print(\"small\")\n");
    assert_eq!(out, vec!["medium"]);
}

#[test]
fn for_range_runtime() {
    let out = run_python("for i in range(5):\n    print(i)\n");
    assert_eq!(out, vec!["0", "1", "2", "3", "4"]);
}

#[test]
fn while_loop_runtime() {
    let out = run_python("i = 0\nwhile i < 3:\n    print(i)\n    i += 1\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_break_runtime() {
    let out = run_python("for i in range(10):\n    if i == 3:\n        break\n    print(i)\n");
    assert_eq!(out, vec!["0", "1", "2"]);
}

#[test]
fn for_continue_runtime() {
    let out = run_python("for i in range(5):\n    if i == 2:\n        continue\n    print(i)\n");
    assert_eq!(out, vec!["0", "1", "3", "4"]);
}

// For/while else

#[test]
fn for_else_no_break() {
    compile_ok("for x in [1, 2, 3]:\n    if x == 5:\n        break\nelse:\n    print('no five')\n");
}

#[test]
fn for_else_with_break() {
    compile_ok("for x in [1, 2, 3]:\n    if x == 2:\n        break\nelse:\n    print('unreachable')\n");
}

#[test]
fn while_else() {
    compile_ok("i = 0\nwhile i < 5:\n    i += 1\nelse:\n    print('done')\n");
}
