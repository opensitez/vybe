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
