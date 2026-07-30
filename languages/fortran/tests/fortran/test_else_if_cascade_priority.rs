use super::helpers::run_prints;

#[test]
fn test_else_if_cascade_priority_resolves_first_match() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority
    integer :: x
    x = 3
    if (x > 4) then
        print *, 1
    else if (x == 3) then
        print *, 2
    else
        print *, 3
    end if
end program test_else_if_cascade_priority
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_else_if_cascade_skips_lower_branches() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_skip
    integer :: x
    x = 2
    if (x > 5) then
        print *, 1
    else if (x >= 2) then
        print *, 2
    else if (x >= 1) then
        print *, 3
    else
        print *, 4
    end if
end program test_else_if_cascade_priority_skip
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_else_if_cascade_default_only() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_default
    integer :: x
    x = -1
    if (x > 5) then
        print *, 1
    else if (x == 3) then
        print *, 2
    else if (x == 2) then
        print *, 3
    else
        print *, 4
    end if
end program test_else_if_cascade_priority_default
"#,
    );

    assert_eq!(out, vec!["4"]);
}

#[test]
fn test_else_if_cascade_with_logical_condition_order() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_logical
    logical :: a, b
    a = .true.
    b = .false.
    if (a .and. b) then
        print *, 1
    else if (a .or. b) then
        print *, 2
    else
        print *, 3
    end if
end program test_else_if_cascade_priority_logical
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_else_if_cascade_short_circuits_after_first_match() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_short_circuit
    integer :: x, y
    x = 9
    y = 0
    if (x > 5) then
        y = 10
    else if (x > 7) then
        y = 20
    else if (x > 8) then
        y = 30
    else
        y = 40
    end if
    print *, y
end program test_else_if_cascade_priority_short_circuit
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_else_if_cascade_nested_false_guard() {
    let out = run_prints(
        r#"
program test_else_if_cascade_nested_false_guard
    real :: v
    v = -1.5
    if (v > 0.0) then
        print *, "pos"
    else if (v < -1.0) then
        print *, "neg-strong"
    else if (abs(v) < 2.0) then
        print *, "small"
    else
        print *, "other"
    end if
end program test_else_if_cascade_nested_false_guard
"#,
    );

    assert_eq!(out, vec!["neg-strong"]);
}

#[test]
fn test_else_if_cascade_repeated_match_after_false_prefix() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_repeated_match
    integer :: x
    x = 2
    if (x == 1) then
        print *, 1
    else if (x == 2) then
        print *, 2
    else if (x == 2) then
        print *, 20
    else
        print *, 3
    end if
end program test_else_if_cascade_priority_repeated_match
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_else_if_cascade_handles_parenthesized_condition() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_parenthesized
    integer :: x
    x = 5
    if ((x + 1) > 10) then
        print *, 1
    else if ((x + 3) == 8) then
        print *, 2
    else
        print *, 3
    end if
end program test_else_if_cascade_priority_parenthesized
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_else_if_cascade_without_final_else_and_all_false() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_no_else
    integer :: x
    x = 0
    if (x > 10) then
        print *, "big"
    else if (x > 5) then
        print *, "medium"
    else if (x > 0) then
        print *, "small"
    end if
    print *, "done"
end program test_else_if_cascade_priority_no_else
"#,
    );

    assert_eq!(out, vec!["done"]);
}

#[test]
fn test_else_if_cascade_character_match_middle_branch() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_char
    character(len=6) :: mode
    mode = "batch "
    if (trim(mode) == "interactive") then
        print *, "interactive"
    else if (trim(mode) == "batch") then
        print *, "batch"
    else if (trim(mode) == "daemon") then
        print *, "daemon"
    else
        print *, "unknown"
    end if
end program test_else_if_cascade_priority_char
"#,
    );

    assert_eq!(out, vec!["batch"]);
}

#[test]
fn test_else_if_cascade_truth_table_ordering() {
    let out = run_prints(
        r#"
program test_else_if_cascade_priority_truth
    logical :: a
    logical :: b
    integer :: c
    a = .true.
    b = .false.
    c = 0
    if (.not. (a .and. b)) then
        c = 10
    else if (a .and. .not. b) then
        c = 20
    else if (b) then
        c = 30
    end if
    print *, c
end program test_else_if_cascade_priority_truth
"#,
    );

    assert_eq!(out, vec!["10"]);
}
