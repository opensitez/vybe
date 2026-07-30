use super::helpers::run_prints;

#[test]
fn test_do_construct_step_semantics_decrements_with_custom_stride() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 10, 2, -4
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_do_construct_step_semantics_single_iteration_when_step_too_large() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 1, 100, 1000
        total = total + 1
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_do_construct_step_semantics_skips_when_direction_mismatch_positive_stride() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 10, 1, 1
        total = total + 1
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_construct_step_semantics_single_step_boundary_is_inclusive() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 1, 1, 1
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_do_construct_step_semantics_descending_three_step() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 9, 1, -3
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["22"]);
}

#[test]
fn test_do_construct_step_semantics_skips_when_direction_mismatch_negative_stride() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10, -2
        total = total + 1
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_do_construct_step_semantics_dynamic_stride_is_bound() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics
    integer :: i
    integer :: step
    integer :: total
    total = 0
    step = 4
    do i = 2, 20, step
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics
"#,
    );

    assert_eq!(out, vec!["56"]);
}

#[test]
fn test_do_construct_step_semantics_negative_start_with_negative_step() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_negative_start
    integer :: i
    integer :: total
    total = 0
    do i = -2, -8, -3
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_negative_start
"#,
    );

    assert_eq!(out, vec!["-15"]);
}

#[test]
fn test_do_construct_step_semantics_step_mutation_is_ignored() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_step_mutation
    integer :: i
    integer :: step
    integer :: total
    step = 1
    total = 0
    do i = 1, 10, step
        if (i == 2) step = 99
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_step_mutation
"#,
    );

    assert_eq!(out, vec!["55"]);
}

#[test]
fn test_do_construct_step_semantics_step_expression() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_step_expression
    integer :: i
    integer :: total
    total = 0
    do i = 1, 10, (1 + 2)
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_step_expression
"#,
    );

    assert_eq!(out, vec!["22"]);
}

#[test]
fn test_do_construct_step_semantics_step_expression_with_intrinsic() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_step_expression_with_intrinsic
    integer :: i
    integer :: total
    total = 0
    do i = 1, 9, mod(13, 2)
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_step_expression_with_intrinsic
"#,
    );

    assert_eq!(out, vec!["25"]);
}

#[test]
fn test_do_construct_step_semantics_step_from_parameter() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_step_from_parameter
    integer, parameter :: stride = 4
    integer :: i
    integer :: total
    total = 0
    do i = 0, 12, stride
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_step_from_parameter
"#,
    );

    assert_eq!(out, vec!["18"]);
}

#[test]
fn test_do_construct_step_semantics_start_and_end_are_expressions() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_expr_bounds
    integer :: a
    integer :: b
    integer :: i
    integer :: total
    a = 2
    b = 5
    total = 0
    do i = a + 1, b * 2, 2
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_expr_bounds
"#,
    );

    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_do_construct_step_semantics_expression_with_negative_step() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_negative_expr_step
    integer :: a
    integer :: i
    integer :: total
    a = 2
    total = 0
    do i = 8, a, -(1 + 1)
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_negative_expr_step
"#,
    );

    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_do_construct_step_semantics_single_iteration_from_expression_bounds() {
    let out = run_prints(
        r#"
program test_do_construct_step_semantics_single_expression_iteration
    integer :: i
    integer :: total
    total = 0
    do i = 2 + 0, 4 - 2, 1 + 1
        total = total + i
    end do
    print *, total
end program test_do_construct_step_semantics_single_expression_iteration
"#,
    );

    assert_eq!(out, vec!["2"]);
}
