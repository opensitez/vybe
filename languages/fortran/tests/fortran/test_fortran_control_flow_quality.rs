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

#[test]
fn control_flow_nested_if_chain_with_elseif() {
    let out = run_prints(
        r#"
program control_flow_nested_if_chain
    integer :: x
    x = 17
    if (x > 20) then
        print *, 1
    else if (x > 10) then
        if (mod(x, 2) == 0) then
            print *, 2
        else
            print *, 3
        end if
    else
        print *, 4
    end if
end program control_flow_nested_if_chain
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn control_flow_if_without_then_false_only_while() {
    let out = run_prints(
        r#"
program control_flow_if_without_then_false_only_while
    integer :: i
    integer :: total
    total = 0
    i = 1
    do while (i < 6)
        if (mod(i, 3) == 0) then
            total = total + i
        end if
        i = i + 1
    end do
    print *, total
end program control_flow_if_without_then_false_only_while
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn control_flow_select_case_character_match() {
    let out = run_prints(
        r#"
program control_flow_select_case_character_match
    character(len=6) :: tag
    tag = "beta  "
    select case (trim(tag))
        case ("alpha")
            print *, 1
        case ("beta")
            print *, 2
        case default
            print *, 3
    end select
end program control_flow_select_case_character_match
"#,
    );

    assert_eq!(out, vec!["2"]);
}

#[test]
fn control_flow_select_case_overlap_prefers_first_clause() {
    let out = run_prints(
        r#"
program control_flow_select_case_overlap_prefers_first_clause
    integer :: n
    n = 5
    select case (n)
        case (1:10)
            print *, 1
        case (5)
            print *, 2
        case default
            print *, 3
    end select
end program control_flow_select_case_overlap_prefers_first_clause
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn control_flow_do_empty_range() {
    let out = run_prints(
        r#"
program control_flow_do_empty_range
    integer :: i
    integer :: total
    total = 0
    do i = 9, 1
        total = total + i
    end do
    print *, total
end program control_flow_do_empty_range
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn control_flow_nested_named_exit() {
    let out = run_prints(
        r#"
program control_flow_nested_named_exit
    integer :: outer_i, inner_i, total
    total = 0
    outer_loop: do outer_i = 1, 5
        do inner_i = 1, 5
            if (outer_i == 4 .and. inner_i == 3) exit outer_loop
            total = total + 1
        end do
    end do outer_loop
    print *, total
end program control_flow_nested_named_exit
"#,
    );

    assert_eq!(out, vec!["17"]);
}

#[test]
fn control_flow_nested_named_cycle_to_outer() {
    let out = run_prints(
        r#"
program control_flow_nested_named_cycle_to_outer
    integer :: outer_i, inner_i, total
    total = 0
    outer_loop: do outer_i = 1, 4
        inner_loop: do inner_i = 1, 5
            if (inner_i == 3) cycle outer_loop
            total = total + 1
        end do inner_loop
    end do outer_loop
    print *, total
end program control_flow_nested_named_cycle_to_outer
"#,
    );

    assert_eq!(out, vec!["8"]);
}

#[test]
fn control_flow_do_while_zero_iterations() {
    let out = run_prints(
        r#"
program control_flow_do_while_zero_iterations
    integer :: i
    integer :: total
    i = 1
    total = 0
    do while (i < 1)
        total = total + 1
        i = i + 1
    end do
    print *, total
end program control_flow_do_while_zero_iterations
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn control_flow_where_elsewhere_mask_chain() {
    let out = run_prints(
        r#"
program control_flow_where_elsewhere_mask_chain
    integer :: a(3)
    integer :: b(3)
    a = [1, -2, 3]
    b = 0
    where (a > 0)
        b = 10
    elsewhere (a < 0)
        b = 20
    elsewhere
        b = 30
    end where
    print *, b(1)
    print *, b(2)
    print *, b(3)
end program control_flow_where_elsewhere_mask_chain
"#,
    );

    assert_eq!(out, vec!["10", "20", "10"]);
}

#[test]
fn control_flow_select_type_class_default() {
    let out = run_prints(
        r#"
program control_flow_select_type_class_default
    class(*), allocatable :: item
    allocate(character(len=5) :: item)
    select type(item)
    type is (integer)
        print *, 1
    type is (real)
        print *, 2
    class default
        print *, 3
    end select
end program control_flow_select_type_class_default
"#,
    );

    assert_eq!(out, vec!["3"]);
}
