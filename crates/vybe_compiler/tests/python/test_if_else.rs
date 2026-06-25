use crate::helpers::{run_python_one, run_python, run_print};

#[test]
fn if_true_branch_prints_yes() {
    let out = run_python("if True:\n    print('yes')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_false_branch_skips_body() {
    let out = run_python("if False:\n    print('yes')\nprint('done')\n");
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_else_selects_if_when_condition_true() {
    let out = run_python("x = 10\nif x > 5:\n    print('big')\nelse:\n    print('small')\n");
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else_selects_else_when_condition_false() {
    let out = run_python("x = 2\nif x > 5:\n    print('big')\nelse:\n    print('small')\n");
    assert_eq!(out, vec!["small"]);
}

#[test]
fn elif_chain_picks_first_matching_branch() {
    let out = run_python(
        "score = 75\nif score >= 90:\n    print('A')\nelif score >= 70:\n    print('B')\nelse:\n    print('C')\n",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn elif_chain_falls_through_to_else() {
    let out = run_python(
        "score = 55\nif score >= 90:\n    print('A')\nelif score >= 70:\n    print('B')\nelse:\n    print('C')\n",
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn elif_chain_stops_at_first_true() {
    let out = run_python(
        "n = 95\nif n >= 90:\n    print('A')\nelif n >= 80:\n    print('B')\nelif n >= 70:\n    print('C')\nelse:\n    print('F')\n",
    );
    assert_eq!(out, vec!["A"]);
}

#[test]
fn elif_chain_middle_branch_wins() {
    let out = run_python(
        "n = 85\nif n >= 90:\n    print('A')\nelif n >= 80:\n    print('B')\nelif n >= 70:\n    print('C')\nelse:\n    print('F')\n",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn elif_chain_lowest_branch() {
    let out = run_python(
        "n = 72\nif n >= 90:\n    print('A')\nelif n >= 80:\n    print('B')\nelif n >= 70:\n    print('C')\nelse:\n    print('F')\n",
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn nested_if_both_conditions_true() {
    let out = run_python("a = 10\nb = 3\nif a > 5:\n    if b < 5:\n        print('inner')\n");
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn nested_if_outer_false_skips_inner() {
    let out = run_python(
        "a = 1\nb = 3\nif a > 5:\n    if b < 5:\n        print('inner')\nprint('end')\n",
    );
    assert_eq!(out, vec!["end"]);
}

#[test]
fn nested_if_inner_false_skips_inner_body() {
    let out = run_python(
        "a = 10\nb = 9\nif a > 5:\n    if b < 5:\n        print('inner')\n    else:\n        print('else-inner')\n",
    );
    assert_eq!(out, vec!["else-inner"]);
}

#[test]
fn nested_if_else_if_chain() {
    let out = run_python(
        "x = 15\nif x < 10:\n    print('low')\nelse:\n    if x < 20:\n        print('mid')\n    else:\n        print('high')\n",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn nested_if_three_levels() {
    let out = run_python(
        "a = 1\nb = 2\nc = 3\nif a:\n    if b:\n        if c:\n            print('deep')\n",
    );
    assert_eq!(out, vec!["deep"]);
}

#[test]
fn truthiness_empty_list_is_falsy() {
    let out = run_python("if []:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["no"]);
}

#[test]
fn truthiness_nonempty_list_is_truthy() {
    let out = run_python("if [1]:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn truthiness_empty_dict_is_falsy() {
    let out = run_python("if {}:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["no"]);
}

#[test]
fn truthiness_nonempty_dict_is_truthy() {
    let out = run_python("if {'a': 1}:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn truthiness_empty_string_is_falsy() {
    let out = run_python("if '':\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["no"]);
}

#[test]
fn truthiness_nonempty_string_is_truthy() {
    let out = run_python("if 'hi':\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn truthiness_zero_is_falsy() {
    let out = run_python("if 0:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["no"]);
}

#[test]
fn truthiness_one_is_truthy() {
    let out = run_python("if 1:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn truthiness_none_is_falsy() {
    let out = run_python("if None:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["no"]);
}

#[test]
fn truthiness_negative_number_is_truthy() {
    let out = run_python("if -1:\n    print('yes')\nelse:\n    print('no')\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn condition_with_and_both_true() {
    let out = run_python("a = 3\nb = 4\nif a > 0 and b > 0:\n    print('both')\n");
    assert_eq!(out, vec!["both"]);
}

#[test]
fn condition_with_and_one_false() {
    let out = run_python("a = 3\nb = -1\nif a > 0 and b > 0:\n    print('both')\nelse:\n    print('not both')\n");
    assert_eq!(out, vec!["not both"]);
}

#[test]
fn condition_with_or_one_true() {
    let out = run_python("a = -1\nb = 4\nif a > 0 or b > 0:\n    print('either')\n");
    assert_eq!(out, vec!["either"]);
}

#[test]
fn condition_with_or_both_false() {
    let out = run_python("a = -1\nb = -2\nif a > 0 or b > 0:\n    print('either')\nelse:\n    print('neither')\n");
    assert_eq!(out, vec!["neither"]);
}

#[test]
fn condition_with_not_inverts_true() {
    let out = run_python("flag = True\nif not flag:\n    print('off')\nelse:\n    print('on')\n");
    assert_eq!(out, vec!["on"]);
}

#[test]
fn condition_with_not_inverts_false() {
    let out = run_python("flag = False\nif not flag:\n    print('off')\nelse:\n    print('on')\n");
    assert_eq!(out, vec!["off"]);
}

#[test]
fn condition_combines_and_or_not() {
    let out = run_python(
        "x = 5\nif (x > 0 and x < 10) or not (x == 3):\n    print('match')\n",
    );
    assert_eq!(out, vec!["match"]);
}

#[test]
fn condition_and_or_precedence_in_if() {
    let out = run_python(
        "a = True\nb = False\nc = True\nif a or b and c:\n    print('yes')\nelse:\n    print('no')\n",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_with_equality_comparison() {
    let out = run_python("name = 'alice'\nif name == 'alice':\n    print('found')\n");
    assert_eq!(out, vec!["found"]);
}

#[test]
fn if_with_inequality_comparison() {
    let out = run_python("n = 7\nif n != 0:\n    print('nonzero')\n");
    assert_eq!(out, vec!["nonzero"]);
}

#[test]
fn if_with_less_equal_chain() {
    let out = run_python("n = 5\nif 1 <= n <= 10:\n    print('in range')\n");
    assert_eq!(out, vec!["in range"]);
}

#[test]
fn if_with_less_equal_chain_outside() {
    let out = run_python("n = 15\nif 1 <= n <= 10:\n    print('in range')\nelse:\n    print('out')\n");
    assert_eq!(out, vec!["out"]);
}

#[test]
fn if_elif_without_final_else() {
    let out = run_python(
        "x = 2\nif x > 10:\n    print('big')\nelif x > 1:\n    print('small')\n",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_elif_no_match_and_no_else() {
    let out = run_python(
        "x = 0\nif x > 10:\n    print('big')\nelif x > 5:\n    print('mid')\nprint('end')\n",
    );
    assert_eq!(out, vec!["end"]);
}

#[test]
fn nested_if_with_elif_in_outer() {
    let out = run_python(
        "x = 5\ny = 2\nif x > 10:\n    print('outer-a')\nelif x > 3:\n    if y < 5:\n        print('inner')\nelse:\n    print('outer-b')\n",
    );
    assert_eq!(out, vec!["inner"]);
}

#[test]
fn if_with_string_membership() {
    let out = run_python("ch = 'a'\nif ch in 'abc':\n    print('member')\n");
    assert_eq!(out, vec!["member"]);
}

#[test]
fn if_with_string_not_in() {
    let out = run_python("ch = 'z'\nif ch not in 'abc':\n    print('absent')\n");
    assert_eq!(out, vec!["absent"]);
}

#[test]
fn if_with_list_membership() {
    let out = run_python("n = 3\nif n in [1, 2, 3]:\n    print('listed')\n");
    assert_eq!(out, vec!["listed"]);
}

#[test]
fn if_is_none_check() {
    let out = run_python("x = None\nif x is None:\n    print('nil')\n");
    assert_eq!(out, vec!["nil"]);
}

#[test]
fn if_is_not_none_check() {
    let out = run_python("x = 42\nif x is not None:\n    print('set')\n");
    assert_eq!(out, vec!["set"]);
}

#[test]
fn if_else_assigns_via_branch() {
    let out = run_python(
        "n = 4\nif n % 2 == 0:\n    label = 'even'\nelse:\n    label = 'odd'\nprint(label)\n",
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn if_elif_else_assigns_grade() {
    let out = run_python(
        "pts = 88\nif pts >= 90:\n    grade = 'A'\nelif pts >= 80:\n    grade = 'B'\nelse:\n    grade = 'C'\nprint(grade)\n",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn run_print_truthiness_in_condition() {
    assert_eq!(run_print("'ok' if 42 else 'no'"), "ok");
}

#[test]
fn run_python_one_if_expression_false_branch() {
    assert_eq!(
        run_python_one("n = 3\nprint('even' if n % 2 == 0 else 'odd')\n"),
        "odd"
    );
}
