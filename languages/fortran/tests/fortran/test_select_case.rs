use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Select case
// ═══════════════════════════════════════════════════════════

#[test]
fn select_case_1() {
    let out = run_prints(
        "program t\ninteger :: x = 1\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["one"]);
}

#[test]
fn select_case_2() {
    let out = run_prints(
        "program t\ninteger :: x = 2\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["two"]);
}

#[test]
fn select_case_default() {
    let out = run_prints(
        "program t\ninteger :: x = 99\nselect case (x)\ncase (1)\nprint *, \"one\"\ncase default\nprint *, \"default\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["default"]);
}

#[test]
fn select_case_no_match_no_default() {
    let out = run_prints(
        "program t\ninteger :: x = 99\nselect case (x)\ncase (1)\nprint *, \"one\"\nend select\nprint *, \"done\"\nend program t\n",
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn select_case_range_match() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 4\nselect case (x)\ncase (1:3)\nprint *, \"small\"\ncase (4:6)\nprint *, \"mid\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn select_case_range_upper_bound_only() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 3\nselect case (x)\ncase (:3)\nprint *, \"low\"\ncase (4:)\nprint *, \"high\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["low"]);
}

#[test]
fn select_case_multiple_labels_one_case() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 2\nselect case (x)\ncase (1, 2, 3)\nprint *, \"abc\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn select_case_empty_range_no_match() {
    let out = run_prints(
        "program t\ninteger :: x\nx = 3\nselect case (x)\ncase (1:2)\nprint *, \"low\"\ncase (4:5)\nprint *, \"high\"\ncase default\nprint *, \"none\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn select_case_negative_range_match() {
    let out = run_prints(
        "program t\ninteger :: x\nx = -2\nselect case (x)\ncase (-5:-3)\nprint *, \"neg\"\ncase (-2:-2)\nprint *, \"minus_two\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["minus_two"]);
}

#[test]
fn select_case_nested_control() {
    let out = run_prints(
        "program t\ninteger :: x\ninteger :: y\nx = 3\ny = 0\nselect case (x)\ncase (1:2)\n y = 1\ncase (3)\n select case (x)\n case (3)\n  y = 3\n case default\n  y = 9\n end select\ncase default\n y = 5\nend select\nprint *, y\nend program t\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn select_case_expression_selector() {
    let out = run_prints(
        "program t\ninteger :: a = 2, b = 5\nselect case (a + b)\ncase (1)\nprint *, \"one\"\ncase (5:8)\nprint *, \"sum-mid\"\ncase (9:)\nprint *, \"sum-high\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["sum-mid"]);
}

#[test]
fn select_case_overlap_precedence() {
    let out = run_prints(
        "program t\ninteger :: n = 7\nselect case (n)\ncase (1:10)\nprint *, \"wide\"\ncase (5:8)\nprint *, \"narrow\"\ncase default\nprint *, \"other\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["wide"]);
}

#[test]
fn select_case_empty_result_when_no_match_no_default() {
    let out = run_prints(
        "program t\ninteger :: n = 99\nselect case (n)\ncase (1)\nprint *, \"one\"\ncase (2)\nprint *, \"two\"\nend select\nend program t\n",
    );
    assert_eq!(out, Vec::<&str>::new());
}

#[test]
fn select_case_default_only_branch() {
    let out = run_prints(
        "program t\ninteger :: n = 99\nselect case (n)\ncase default\nprint *, \"fallback\"\nend select\nend program t\n",
    );
    assert_eq!(out, vec!["fallback"]);
}

#[test]
fn select_case_logical_selector() {
    let out = run_prints(
        r#"
program t
logical :: active
active = .true.
select case (active)
case (.true.)
print *, 'on'
case (.false.)
print *, 'off'
case default
print *, 'default'
end select
end program t
"#,
    );
    assert_eq!(out, vec!["on"]);
}

#[test]
fn select_case_character_selector_with_range_and_default() {
    let out = run_prints(
        r#"
program t
character(len=1) :: grade
grade = 'B'
select case (grade)
case ('A':'C')
print *, 'good'
case ('D':'F')
print *, 'bad'
case default
print *, 'other'
end select
end program t
"#,
    );
    assert_eq!(out, vec!["good"]);
}

#[test]
fn select_case_multiple_case_items_overlap_by_list_then_range() {
    let out = run_prints(
        r#"
program t
integer :: n
n = 4
select case (n)
case (3, 4, 5)
print *, 'list'
case (1:10)
print *, 'range'
end select
end program t
"#,
    );
    assert_eq!(out, vec!["list"]);
}

#[test]
fn select_case_in_loop() {
    let out = run_prints(
        r#"
program t
integer :: i
do i = 1, 4
 select case (mod(i, 2))
 case (0)
  print *, 'even'
 case (1)
  print *, 'odd'
 end select
end do
end program t
"#,
    );
    assert_eq!(out, vec!["odd", "even", "odd", "even"]);
}
