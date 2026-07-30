use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: If/then/else blocks
// ═══════════════════════════════════════════════════════════

#[test]
fn if_true() {
    let out = run_prints("program t\nif (1 > 0) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_false() {
    let out = run_prints(
        "program t\nif (0 > 1) then\nprint *, \"no\"\nend if\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_else_true() {
    let out = run_prints(
        "program t\nif (5 > 3) then\nprint *, \"big\"\nelse\nprint *, \"small\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_else_false() {
    let out = run_prints(
        "program t\nif (1 > 5) then\nprint *, \"big\"\nelse\nprint *, \"small\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_elseif() {
    let out = run_prints(
        "program t\ninteger :: x = 2\nif (x == 1) then\nprint *, \"one\"\nelse if (x == 2) then\nprint *, \"two\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn if_elseif_default() {
    let out = run_prints(
        "program t\ninteger :: x = 99\nif (x == 1) then\nprint *, \"one\"\nelse if (x == 2) then\nprint *, \"two\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn if_gt() {
    let out = run_prints("program t\nif (5 > 3) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_lt() {
    let out = run_prints("program t\nif (3 < 5) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_eq() {
    let out =
        run_prints("program t\nif (3 == 3) then\nprint *, \"equal\"\nend if\nend program t\n");
    assert_eq!(out, vec!["equal"]);
}

#[test]
fn if_ne() {
    let out =
        run_prints("program t\nif (3 /= 4) then\nprint *, \"not equal\"\nend if\nend program t\n");
    assert_eq!(out, vec!["not equal"]);
}

#[test]
fn if_ge() {
    let out = run_prints("program t\nif (5 >= 5) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_le() {
    let out = run_prints("program t\nif (3 <= 3) then\nprint *, \"yes\"\nend if\nend program t\n");
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_with_variable() {
    let out = run_prints(
        "program t\ninteger :: score = 85\nif (score >= 90) then\nprint *, \"A\"\nelse if (score >= 80) then\nprint *, \"B\"\nelse\nprint *, \"C\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn if_logical_and() {
    let out = run_prints(
        "program t\nif (1 > 0 .and. 2 > 1) then\nprint *, \"both\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["both"]);
}

#[test]
fn if_logical_or() {
    let out = run_prints(
        "program t\nif (1 > 5 .or. 2 > 1) then\nprint *, \"either\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["either"]);
}

#[test]
fn if_logical_not() {
    let out = run_prints(
        "program t\nif (.not. (1 > 5)) then\nprint *, \"negated\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["negated"]);
}

#[test]
fn nested_if() {
    let out = run_prints(
        "program t\nif (1 > 0) then\nif (2 > 1) then\nprint *, \"nested\"\nend if\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["nested"]);
}

#[test]
fn single_line_if_print() {
    let out = run_prints("program t\nif (1 > 0) print *, \"inline\"\nend program t\n");
    assert_eq!(out, vec!["inline"]);
}

#[test]
fn if_after_assignment() {
    let out = run_prints(
        "program t\ninteger :: x = 10\nx = x + 5\nif (x == 15) then\nprint *, \"correct\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["correct"]);
}

#[test]
fn if_multiple_statements_in_body() {
    let out = run_prints(
        "program t\nif (1 > 0) then\nprint *, \"a\"\nprint *, \"b\"\nprint *, \"c\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn single_line_if_false_path() {
    let out = run_prints("program t\nif (1 == 0) print *, \"no\"\nprint *, \"ok\"\nend program t\n");
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn if_else_multiple_statements_false_branch() {
    let out = run_prints(
        "program t\nif (1 == 0) then\nprint *, \"a\"\nprint *, \"b\"\nelse\nprint *, \"c\"\nprint *, \"d\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["c", "d"]);
}

#[test]
fn if_logical_parentheses_and_not() {
    let out = run_prints(
        "program t\nif ((1 > 0) .and. .not. (2 < 3)) then\nprint *, \"true\"\nelse\nprint *, \"false\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn if_with_logical_variable_true() {
    let out = run_prints(
        "program t\nlogical :: cond\ncond = .true.\nif (cond) then\nprint *, \"yes\"\nelse\nprint *, \"no\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_with_logical_variable_false() {
    let out = run_prints(
        "program t\nlogical :: cond\ncond = .false.\nif (cond) then\nprint *, \"yes\"\nelse\nprint *, \"no\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["no"]);
}

#[test]
fn if_without_final_else_no_branch_taken() {
    let out = run_prints(
        "program t\nif (1 == 0) then\nprint *, \"should not print\"\nend if\nprint *, \"after\"\nend program t\n",
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn if_elseif_chain_missing_final_else() {
    let out = run_prints(
        "program t\ninteger :: x = 2\nif (x == 3) then\nprint *, \"three\"\nelse if (x == 2) then\nprint *, \"two\"\nelse if (x == 1) then\nprint *, \"one\"\nend if\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["two", "done"]);
}

#[test]
fn if_elseif_chain_all_false_no_else() {
    let out = run_prints(
        "program t\ninteger :: x = 0\nif (x == 3) then\nprint *, \"three\"\nelse if (x == 2) then\nprint *, \"two\"\nelse if (x == 1) then\nprint *, \"one\"\nend if\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_logical_eqv_and_neqv() {
    let out = run_prints(
        "program t\nlogical :: a = .true., b = .false.\nif (a .eqv. b) then\nprint *, \"eqv\"\nend if\nif ((a .neqv. b)) then\nprint *, \"neqv\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["neqv"]);
}

#[test]
fn if_with_arithmetic_expression_condition() {
    let out = run_prints(
        "program t\ninteger :: x = 2, y = 3\nif ((x + y) * 2 > 10) then\nprint *, \"big\"\nelse\nprint *, \"small\"\nend if\nif (x * y + 1 == 7) then\nprint *, \"seven\"\nelse\nprint *, \"not seven\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["small", "seven"]);
}

#[test]
fn if_with_logical_grouping_and_operators() {
    let out = run_prints(
        "program t\nlogical :: a = .true., b = .false., c = .true.\nif ((a .and. b) .or. ((.not. a) .and. c)) then\nprint *, \"branch1\"\nelse if ((a .and. c) .and. (.not. b)) then\nprint *, \"branch2\"\nelse\nprint *, \"branch3\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["branch2"]);
}

#[test]
fn nested_if_with_deep_else_if() {
    let out = run_prints(
        "program t\ninteger :: x = 4\nif (x > 0) then\nif (x > 3) then\nif (x > 5) then\nprint *, \"high\"\nelse if (x == 4) then\nprint *, \"mid\"\nelse\nprint *, \"low\"\nend if\nelse\nprint *, \"inner else\"\nend if\nelse\nprint *, \"outer else\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn if_in_loop_body() {
    let out = run_prints(
        "program t\ninteger :: i\ninteger :: n = 0\ndo i = 1, 3\nif (mod(i, 2) == 1) then\nn = n + 1\nend if\nend do\nprint *, n\nend program t\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn if_with_end_if_token() {
    let out = run_prints(
        "program t\nif (1 > 0) then\nprint *, \"ok\"\nelse\nprint *, \"bad\"\nendif\nend program t\n",
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn if_else_if_chain_uses_spacing_form() {
    let out = run_prints(
        "program t\ninteger :: x = 3\nif (x == 1) then\nprint *, \"one\"\nelse if (x == 2) then\nprint *, \"two\"\nelse if (x == 3) then\nprint *, \"three\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["three"]);
}

#[test]
fn if_with_intrinsic_in_condition() {
    let out = run_prints(
        "program t\ninteger :: x = 2\nif (mod(x, 2) == 0) then\nprint *, \"even\"\nelse\nprint *, \"odd\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["even"]);
}

#[test]
fn if_condition_with_min() {
    let out = run_prints(
        "program t\ninteger :: x = 4\nif (min(x, 10) == x) then\nprint *, \"small\"\nelse\nprint *, \"large\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn if_then_without_true_statements() {
    let out = run_prints(
        "program t\nif (.false.) then\nprint *, \"no\"\nelse\nprint *, \"yes\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn if_without_false_branch() {
    let out = run_prints(
        "program t\ninteger :: x = 0\nif (x > 0) then\nprint *, \"positive\"\nend if\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn if_scoped_assignment_updates_once() {
    let out = run_prints(
        "program t\ninteger :: x, y\nx = 1\ny = 0\nif (x == 1) then\n  y = 10\nelse\n  y = 20\nend if\nif (y == 10) then\n  print *, \"ten\"\nelse\n  print *, \"other\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["ten"]);
}

#[test]
fn if_elseif_no_match_and_trailing_statements() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 99\nif (x == 1) then\n  print *, 1\nelse if (x == 2) then\n  print *, 2\nelse if (x == 3) then\n  print *, 3\nend if\nprint *, 9\n",
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn if_logical_chain_true_mid_branch() {
    let out = run_prints(
        "program t\ninteger :: x\nlogical :: ok\nx = 5\nok = .false.\nif (x < 0) then\n  print *, \"neg\"\nelse if (x > 0 .and. .not. ok) then\n  print *, \"pos\"\nelse if (x == 5) then\n  print *, \"five\"\nelse\n  print *, \"other\"\nend if\n",
    );
    assert_eq!(out, vec!["pos"]);
}

#[test]
fn if_character_trimmed_match() {
    let out = run_prints(
        "program t\ncharacter(len=6) :: word\nword = 'Alpha '\nif (trim(word) == 'Alpha') then\n  print *, 'match'\nelse\n  print *, 'nomatch'\nend if\n",
    );
    assert_eq!(out, vec!["match"]);
}

#[test]
fn if_nested_in_else_branch() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 1\nif (x == 0) then\n  print *, \"outer\"\nelse\n  if (x == 1) then\n    print *, \"inner-one\"\n  else\n    print *, \"inner-other\"\n  end if\nend if\n",
    );
    assert_eq!(out, vec!["inner-one"]);
}

#[test]
fn if_condition_precedence_with_parentheses() {
    let out = run_prints(
        "program t\nlogical :: a = .true.\nlogical :: b = .false.\nlogical :: c = .true.\nif (a .and. b .or. .not. c) then\nprint *, \"wrong\"\nelse if ((a .and. (.not. b)) .or. c) then\nprint *, \"right\"\nelse\nprint *, \"other\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["right"]);
}

#[test]
fn if_with_arithmetic_and_character_conditions() {
    let out = run_prints(
        "program t\ninteger :: i\ncharacter(len=4) :: label\nlabel = 'done'\ni = 2\nif ((i + 1 == 3) .and. (trim(label) == 'done')) then\nprint *, \"pass\"\nelse\nprint *, \"fail\"\nend if\nend program t\n",
    );
    assert_eq!(out, vec!["pass"]);
}
