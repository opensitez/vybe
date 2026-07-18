use super::helpers::run_prints;

#[test]
fn control_flow_if_true_and_false() {
    let out = run_prints(
        r#"
program control_flow_if_true_and_false
    integer :: value
    if (.true.) then
        value = 1
    else
        value = 2
    end if
    print *, value
end program control_flow_if_true_and_false
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn control_flow_if_false_block() {
    let out = run_prints(
        r#"
program control_flow_if_false_block
    integer :: value
    if (.false.) then
        value = 1
    else
        value = 2
    end if
    print *, value
end program control_flow_if_false_block
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn control_flow_if_elseif_chain() {
    let out = run_prints(
        r#"
program control_flow_if_elseif_chain
    integer :: score
    score = 77
    if (score >= 90) then
        print *, 1
    else if (score >= 80) then
        print *, 2
    else if (score >= 70) then
        print *, 3
    else
        print *, 4
    end if
end program control_flow_if_elseif_chain
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn control_flow_logical_precedence() {
    let out = run_prints(
        r#"
program control_flow_logical_precedence
    logical :: a, b, c
    a = .true.
    b = .false.
    c = .true.
    if (a .and. b .or. c) then
        print *, 1
    else
        print *, 0
    end if
end program control_flow_logical_precedence
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn control_flow_select_case_exact() {
    let out = run_prints(
        r#"
program control_flow_select_case_exact
    integer :: mode
    mode = 2
    select case (mode)
        case (1)
            print *, 10
        case (2)
            print *, 20
        case default
            print *, 30
    end select
end program control_flow_select_case_exact
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn control_flow_select_case_range() {
    let out = run_prints(
        r#"
program control_flow_select_case_range
    integer :: score
    score = 67
    select case (score)
        case (0:50)
            print *, 1
        case (51:80)
            print *, 2
        case (81:)
            print *, 3
    end select
end program control_flow_select_case_range
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn control_flow_while_accumulator() {
    let out = run_prints(
        r#"
program control_flow_while_accumulator
    integer :: n
    integer :: total
    n = 0
    total = 0
    do while (n < 5)
        n = n + 1
        total = total + n
    end do
    print *, total
end program control_flow_while_accumulator
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn control_flow_do_ascending() {
    let out = run_prints(
        r#"
program control_flow_do_ascending
    integer :: i
    integer :: total
    total = 0
    do i = 1, 5
        total = total + i
    end do
    print *, total
end program control_flow_do_ascending
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn control_flow_do_descending_step() {
    let out = run_prints(
        r#"
program control_flow_do_descending_step
    integer :: i
    integer :: total
    total = 0
    do i = 10, 2, -2
        total = total + i
    end do
    print *, total
end program control_flow_do_descending_step
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn control_flow_do_with_cycle() {
    let out = run_prints(
        r#"
program control_flow_do_with_cycle
    integer :: i
    integer :: total
    total = 0
    do i = 1, 8
        if (mod(i, 3) == 0) cycle
        total = total + i
    end do
    print *, total
end program control_flow_do_with_cycle
"#,
    );
    assert_eq!(out, vec!["26"]);
}

#[test]
fn control_flow_do_with_exit() {
    let out = run_prints(
        r#"
program control_flow_do_with_exit
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (i > 6) exit
        total = total + i
    end do
    print *, total
end program control_flow_do_with_exit
"#,
    );
    assert_eq!(out, vec!["21"]);
}

#[test]
fn control_flow_named_loop() {
    let out = run_prints(
        r#"
program control_flow_named_loop
    integer :: i
    integer :: total
    total = 0
    outer: do i = 1, 20
        if (i > 4) exit outer
        total = total + i
    end do outer
    print *, total
end program control_flow_named_loop
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn control_flow_nested_if_and_loop() {
    let out = run_prints(
        r#"
program control_flow_nested_if_and_loop
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (mod(i, 2) == 0) then
            if (i <= 6) total = total + i
        end if
    end do
    print *, total
end program control_flow_nested_if_and_loop
"#,
    );
    assert_eq!(out, vec!["12"]);
}

#[test]
fn control_flow_mixed_assignment_flow() {
    let out = run_prints(
        r#"
program control_flow_mixed_assignment_flow
    integer :: x
    integer :: y
    x = 5
    y = 0
    if (x > 3) then
        y = x * 2
    else
        y = x + 1
    end if
    do while (y < 20)
        y = y + 3
    end do
    print *, y
end program control_flow_mixed_assignment_flow
"#,
    );
    assert_eq!(out, vec!["20"]);
}
