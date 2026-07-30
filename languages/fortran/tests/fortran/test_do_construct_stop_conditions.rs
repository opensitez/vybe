use super::helpers::run_prints;

#[test]
fn test_do_construct_stop_conditions_exit_at_threshold() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (i > 4) exit
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_do_construct_stop_conditions_exit_on_first_iteration() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (i == 1) exit
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_construct_stop_conditions_exit_no_stop() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 4
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_do_construct_stop_conditions_cycle_skips() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 6
        if (mod(i, 2) == 0) cycle
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn test_do_construct_stop_conditions_nested_named_exit_outer() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 10
        do inner = 1, 10
            total = total + 1
            if (outer == 2 .and. inner == 3) exit outer_loop
        end do
    end do outer_loop
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["13"]);
}

#[test]
fn test_do_construct_stop_conditions_named_cycle_ignored_outer() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        inner_loop: do inner = 1, 4
            if (inner == 3) cycle outer_loop
            total = total + 1
        end do inner_loop
    end do outer_loop
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["8"]);
}

#[test]
fn test_do_construct_stop_conditions_exit_with_logical_guard() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    logical :: active
    total = 0
    active = .true.
    do i = 1, 10
        if (i == 4) active = .false.
        if (.not. active) exit
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_do_construct_stop_conditions_exit_when_sum_exceeds_limit() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        total = total + i
        if (total > 8) exit
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_do_construct_stop_conditions_cycle_then_exit() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10
        if (mod(i, 3) == 0) cycle
        total = total + i
        if (total > 12) exit
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["19"]);
}

#[test]
fn test_do_construct_stop_conditions_zero_step_invalid_empty() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 1, 5, 0
        total = total + 1
        exit
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_do_construct_stop_conditions_descending_full_range() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 9, 1, -2
        total = total + i
        if (i <= 3) exit
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["24"]);
}

#[test]
fn test_do_construct_stop_conditions_nested_stop_and_cycle_mix() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: j
    integer :: total
    total = 0
    do i = 1, 4
        do j = 1, 4
            if (j == 3) cycle
            total = total + 1
            if (j == 4) exit
        end do
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["12"]);
}

#[test]
fn test_do_construct_stop_conditions_inner_exit_outer_cycle() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: outer, inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        inner_loop: do inner = 1, 4
            if (outer == 2 .and. inner == 2) cycle outer_loop
            if (inner == 4) exit
            total = total + 1
        end do inner_loop
    end do outer_loop
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_do_construct_stop_conditions_mutated_bound_does_not_extend_loop() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions_mutated_bound
    integer :: i
    integer :: total
    integer :: stop
    stop = 4
    total = 0
    do i = 1, stop
        if (i == 2) stop = 10
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions_mutated_bound
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_do_construct_stop_conditions_do_while_exit() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 8)
        i = i + 1
        if (i == 6) exit
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_do_construct_stop_conditions_do_while_cycle() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 6)
        i = i + 1
        if (mod(i, 2) == 0) cycle
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn test_do_construct_stop_conditions_descending_with_negative_cycle_to_next() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    do i = 6, 1, -1
        if (mod(i, 2) == 0) cycle
        total = total + i
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["9"]);
}

#[test]
fn test_do_construct_stop_conditions_zero_step_nested_no_progress_guarded_by_if() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: i
    integer :: total
    total = 0
    i = 1
    do while (i < 2)
        if (i == 1) then
            i = i + 1
            cycle
        end if
        total = total + 1
    end do
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_construct_stop_conditions_named_exit_prefers_target() {
    let out = run_prints(
        r#"
program test_do_construct_stop_conditions
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        do inner = 1, 4
            if (outer == 3 .and. inner == 2) exit outer_loop
            total = total + 1
        end do
    end do outer_loop
    print *, total
end program test_do_construct_stop_conditions
"#,
    );

    assert_eq!(out, vec!["8"]);
}
