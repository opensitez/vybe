use super::helpers::run_prints;

// ═══════════════════════════════════════════════════════════
// Fortran: Control flow — if/then/else, do, select case
// ═══════════════════════════════════════════════════════════

#[test]
fn if_then_else() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 5
    if (x > 3) then
        print *, "big"
    else
        print *, "small"
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn if_elseif_chain() {
    let out = run_prints(
        r#"
program test
    integer :: score
    score = 75
    if (score >= 90) then
        print *, "A"
    else if (score >= 80) then
        print *, "B"
    else if (score >= 70) then
        print *, "C"
    else
        print *, "F"
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["C"]);
}

#[test]
fn single_line_if() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 5
    if (x > 3) print *, "big"
end program test
"#,
    );
    assert_eq!(out, vec!["big"]);
}

#[test]
fn do_loop() {
    let out = run_prints(
        r#"
program test
    integer :: i, sum
    sum = 0
    do i = 1, 5
        sum = sum + i
    end do
    print *, sum
end program test
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn do_loop_with_step() {
    let out = run_prints(
        r#"
program test
    integer :: i, sum
    sum = 0
    do i = 0, 10, 2
        sum = sum + i
    end do
    print *, sum
end program test
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn do_while_loop() {
    let out = run_prints(
        r#"
program test
    integer :: i
    i = 0
    do while (i < 3)
        i = i + 1
    end do
    print *, i
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn nested_do_loops() {
    let out = run_prints(
        r#"
program test
    integer :: i, j, count
    count = 0
    do i = 1, 3
        do j = 1, 4
            count = count + 1
        end do
    end do
    print *, count
end program test
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn exit_do_loop() {
    let out = run_prints(
        r#"
program test
    integer :: i, sum
    sum = 0
    do i = 1, 100
        if (i > 5) exit
        sum = sum + i
    end do
    print *, sum
end program test
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn cycle_do_loop() {
    let out = run_prints(
        r#"
program test
    integer :: i, sum
    sum = 0
    do i = 1, 10
        if (mod(i, 2) /= 0) cycle
        sum = sum + i
    end do
    print *, sum
end program test
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn select_case_basic() {
    let out = run_prints(
        r#"
program test
    integer :: day
    day = 3
    select case (day)
        case (1)
            print *, "Monday"
        case (2)
            print *, "Tuesday"
        case (3)
            print *, "Wednesday"
        case default
            print *, "Other"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["Wednesday"]);
}

#[test]
fn select_case_default() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 99
    select case (x)
        case (1)
            print *, "one"
        case default
            print *, "other"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["other"]);
}

#[test]
fn logical_operators() {
    let out = run_prints(
        r#"
program test
    logical :: a, b
    a = .true.
    b = .false.
    if (a .and. .not. b) then
        print *, "yes"
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn comparison_operators() {
    let out = run_prints(
        r#"
program test
    if (1 < 2) print *, "lt"
    if (2 > 1) print *, "gt"
    if (1 <= 1) print *, "le"
    if (1 >= 1) print *, "ge"
    if (1 == 1) print *, "eq"
    if (1 /= 2) print *, "ne"
end program test
"#,
    );
    assert_eq!(out, vec!["lt", "gt", "le", "ge", "eq", "ne"]);
}

#[test]
fn if_then_else_false_path() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 2
    if (x > 3) then
        print *, "big"
    else
        print *, "small"
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["small"]);
}

#[test]
fn single_line_if_false_path() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 2
    if (x > 3) print *, "big"
    print *, "after"
end program test
"#,
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn do_loop_downward_step() {
    let out = run_prints(
        r#"
program test
    integer :: i, sum
    sum = 0
    do i = 5, 1, -2
        sum = sum + i
    end do
    print *, sum
end program test
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn do_loop_zero_iterations_with_named_cycle_target() {
    let out = run_prints(
        r#"
program test
    integer :: i
    integer :: acc
    acc = 0

    outer: do while (.false.)
        acc = acc + 1
        cycle outer
    end do outer

    print *, acc
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn do_while_zero_iterations() {
    let out = run_prints(
        r#"
program test
    integer :: i
    i = 0
    do while (i < 0)
        i = i + 1
    end do
    print *, i
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn named_do_cycle_outer() {
    let out = run_prints(
        r#"
program test
    integer :: i
    integer :: acc
    acc = 0

    outer: do i = 1, 4
        if (mod(i, 2) == 1) cycle outer
        acc = acc + i
    end do outer

    print *, acc
end program test
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn named_do_exit() {
    let out = run_prints(
        r#"
program test
    integer :: i
    integer :: j
    outer: do i = 1, 5
        inner: do j = 1, 5
            if (i == 2 .and. j == 3) exit outer
        end do inner
    end do outer
    print *, i
end program test
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn if_single_branch_no_else_true_false_paths() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 0
    if (x > 0) then
        print *, 'positive'
    end if
    if (x <= 0) then
        print *, 'nonpositive'
    end if
end program test
"#,
    );
    assert_eq!(out, vec!["nonpositive"]);
}

#[test]
fn select_case_ranges() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 4
    select case (x)
        case (1:3)
            print *, "low"
        case (4:6)
            print *, "mid"
        case default
            print *, "high"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["mid"]);
}

#[test]
fn select_case_open_ranges() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 2
    select case (x)
        case (:4)
            print *, "low"
        case (5:)
            print *, "high"
        case default
            print *, "mid"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["low"]);
}

#[test]
fn select_case_open_high_range() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 9
    select case (x)
        case (:3)
            print *, "low"
        case (4,6)
            print *, "middle"
        case (7:)
            print *, "high"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["high"]);
}

#[test]
fn select_case_character_match() {
    let out = run_prints(
        r#"
program test
    character(len=5) :: mode
    mode = "beta"
    select case (trim(mode))
        case ("alpha", "beta")
            print *, "ok"
        case ("gamma")
            print *, "skip"
        case default
            print *, "other"
    end select
end program test
"#,
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn select_case_no_default_no_match() {
    let out = run_prints(
        r#"
program test
    integer :: x
    x = 99
    select case (x)
    case (1)
        print *, "one"
    case (2)
        print *, "two"
    end select
    print *, "after"
end program test
"#,
    );
    assert_eq!(out, vec!["after"]);
}

#[test]
fn if_elseif_chain_without_final_else_no_match() {
    let out = run_prints(
        r#"
program test
    integer :: score
    score = 50
    if (score >= 90) then
        print *, "a"
    else if (score >= 80) then
        print *, "b"
    else if (score >= 70) then
        print *, "c"
    end if
    print *, "done"
end program test
"#,
    );
    assert_eq!(out, vec!["done"]);
}

#[test]
fn do_loop_start_mutation_ignored_for_iteration_bound() {
    let out = run_prints(
        r#"
program test
    integer :: i, n, s
    n = 2
    s = 0
    do i = 1, n
        if (i == 1) n = 99
        s = s + i
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn do_loop_zero_iterations_with_positive_step_on_descending_bounds() {
    let out = run_prints(
        r#"
program test
    integer :: i
    integer :: s
    s = 0
    do i = 10, 1, 1
        s = s + 1
    end do
    print *, s
end program test
"#,
    );
    assert_eq!(out, vec!["0"]);
}
