use super::helpers::run_prints;

#[test]
fn test_do_while_progress_guarantees_monotonic_counter() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 5)
        i = i + 1
        total = total + i
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees
"#,
    );

    assert_eq!(out, vec!["5", "15"]);
}

#[test]
fn test_do_while_progress_guarantees_terminates_on_condition_update() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_terminates_on_condition_update
    integer :: i
    integer :: total
    i = 1
    total = 0
    do while (i < 20)
        if (mod(i, 2) == 0) then
            total = total + i
        end if
        i = i + 1
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees_terminates_on_condition_update
"#,
    );

    assert_eq!(out, vec!["20", "90"]);
}

#[test]
fn test_do_while_progress_guarantees_exit_resets_flag() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_exit_resets_flag
    integer :: n
    logical :: keep
    n = 0
    keep = .true.
    do while (keep)
        n = n + 1
        if (n == 3) keep = .false.
    end do
    print *, n
end program test_do_while_progress_guarantees_exit_resets_flag
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_do_while_zero_iteration_when_initial_false() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_zero_iteration
    integer :: i
    i = 0
    do while (.false.)
        i = i + 1
    end do
    print *, i
end program test_do_while_progress_guarantees_zero_iteration
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_while_two_step_progress_guarantees() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_two_step
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 9)
        i = i + 2
        total = total + i
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees_two_step
"#,
    );

    assert_eq!(out, vec!["10", "30"]);
}

#[test]
fn test_do_while_progress_control_set_in_branch() {
    let out = run_prints(
        r#"
    program test_do_while_progress_control_set_in_branch
    integer :: n
    logical :: ok
    n = 0
    ok = .true.
    do while (ok)
        n = n + 1
        if (n >= 3) then
            ok = .false.
        else
            n = n + 1
        end if
    end do
    print *, n
end program test_do_while_progress_control_set_in_branch
"#,
    );

    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_do_while_progress_guarantees_nested_do_cycle_progress() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_nested_do_cycle_progress
    integer :: outer
    integer :: inner
    integer :: count
    outer = 0
    count = 0
    do while (outer < 4)
        outer = outer + 1
        inner = 0
        do while (inner < 4)
            inner = inner + 1
            if (inner == 2) cycle
            count = count + 1
        end do
    end do
    print *, outer
    print *, count
end program test_do_while_progress_guarantees_nested_do_cycle_progress
"#,
    );
    assert_eq!(out, vec!["4", "12"]);
}

#[test]
fn test_do_while_progress_guarantees_condition_flip_break() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_condition_flip_break
    integer :: n
    logical :: active
    n = 0
    active = .true.
    do while (active)
        n = n + 1
        if (n == 6) active = .false.
        if (mod(n, 2) == 1) cycle
        if (n > 8) active = .false.
    end do
    print *, n
end program test_do_while_progress_guarantees_condition_flip_break
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_do_while_progress_guarantees_cycle_without_progress_on_counter() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_cycle_without_progress_on_counter
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 3)
        if (mod(total + 1, 2) == 0) then
            total = total + 1
            cycle
        end if
        i = i + 1
        total = total + 1
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees_cycle_without_progress_on_counter
"#,
    );

    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn test_do_while_progress_guarantees_guarded_exit_respects_condition_flow() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_guarded_exit_respects_condition_flow
    integer :: i
    integer :: total
    i = 0
    total = 0
    do while (i < 10)
        i = i + 1
        if (i == 4) then
            exit
        end if
        total = total + i
    end do
    print *, i
    print *, total
end program test_do_while_progress_guarantees_guarded_exit_respects_condition_flow
"#,
    );

    assert_eq!(out, vec!["4", "6"]);
}

#[test]
fn test_do_while_progress_guarantees_nested_progress_after_cycle() {
    let out = run_prints(
        r#"
program test_do_while_progress_guarantees_nested_progress_after_cycle
    integer :: outer
    integer :: inner
    integer :: total
    outer = 0
    total = 0
    do while (outer < 3)
        outer = outer + 1
        inner = 0
        do while (inner < 4)
            inner = inner + 1
            if (inner == 2) cycle
            if (inner == 3) exit
            total = total + 1
        end do
    end do
    print *, outer
    print *, total
end program test_do_while_progress_guarantees_nested_progress_after_cycle
"#,
    );

    assert_eq!(out, vec!["3", "3"]);
}
