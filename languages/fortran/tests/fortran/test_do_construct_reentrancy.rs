use super::helpers::run_prints;

#[test]
fn test_do_construct_reentrancy_nested_accumulators() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy
    integer :: outer, inner, total
    total = 0
    do outer = 1, 3
        do inner = 1, 2
            total = total + outer * inner
        end do
    end do
    print *, total
end program test_do_construct_reentrancy
"#,
    );

    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_do_construct_reentrancy_nested_early_exit() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_early_exit
    integer :: outer, inner, total
    total = 0
    do outer = 1, 4
        do inner = 1, 5
            total = total + 1
            if (outer == 3 .and. inner == 2) exit
        end do
    end do
    print *, total
end program test_do_construct_reentrancy_early_exit
"#,
    );

    assert_eq!(out, vec!["17"]);
}

#[test]
fn test_do_construct_reentrancy_descending_inner_loop() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_descending
    integer :: outer, inner, total
    total = 0
    do outer = 1, 2
        do inner = 5, 1, -2
            total = total + inner
        end do
    end do
    print *, total
end program test_do_construct_reentrancy_descending
"#,
    );

    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_do_construct_reentrancy_bare_loop_with_exit() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_bare_loop
    integer :: i, total
    i = 0
    total = 0
    do
        i = i + 1
        if (i == 4) exit
        total = total + i
    end do
    print *, i
    print *, total
end program test_do_construct_reentrancy_bare_loop
"#,
    );

    assert_eq!(out, vec!["4", "6"]);
}

#[test]
fn test_do_construct_reentrancy_named_outer_cycle_skips_outer_iteration() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_named_outer_cycle
    integer :: outer, inner, total
    total = 0
    outer_loop: do outer = 1, 4
        do inner = 1, 3
            if (inner == 2 .and. outer == 2) cycle outer_loop
            total = total + 1
        end do
    end do outer_loop
    print *, outer
    print *, total
end program test_do_construct_reentrancy_named_outer_cycle
"#,
    );

    assert_eq!(out, vec!["5", "10"]);
}

#[test]
fn test_do_construct_reentrancy_exit_inner_only_continues_outer() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_exit_inner_only
    integer :: outer, inner, total
    total = 0
    do outer = 1, 3
        do inner = 1, 5
            if (inner == 3) exit
            total = total + outer + inner
        end do
    end do
    print *, total
end program test_do_construct_reentrancy_exit_inner_only
"#,
    );

    assert_eq!(out, vec!["15"]);
}

#[test]
fn test_do_construct_reentrancy_named_inner_exit_only() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_named_inner_exit
    integer :: outer, inner, total
    total = 0
    outer_loop: do outer = 1, 3
        inner_loop: do inner = 1, 4
            if (inner == 2) exit inner_loop
            total = total + outer
        end do inner_loop
    end do outer_loop
    print *, total
end program test_do_construct_reentrancy_named_inner_exit
"#,
    );

    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_do_construct_reentrancy_empty_trip_no_iteration() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_empty_trip
    integer :: i, sum
    sum = 0
    do i = 1, 5, -1
        sum = sum + i
    end do
    print *, sum
end program test_do_construct_reentrancy_empty_trip
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_construct_reentrancy_do_with_step_two_and_cycle() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_step_two
    integer :: i, total
    total = 0
    do i = 0, 10, 2
        if (i == 6) cycle
        total = total + i
    end do
    print *, total
end program test_do_construct_reentrancy_step_two
"#,
    );

    assert_eq!(out, vec!["24"]);
}

#[test]
fn test_do_construct_reentrancy_named_cycle_advances_outer_index() {
    let out = run_prints(
        r#"
program test_do_construct_reentrancy_named_cycle
    integer :: outer
    integer :: inner
    integer :: total
    total = 0
    outer_loop: do outer = 1, 4
        do inner = 1, 4
            if (mod(inner, 2) == 0) cycle outer_loop
            total = total + outer
        end do
    end do outer_loop
    print *, total
    print *, outer
    print *, inner
end program test_do_construct_reentrancy_named_cycle
"#,
    );
    assert_eq!(out, vec!["6", "5", "2"]);
}
