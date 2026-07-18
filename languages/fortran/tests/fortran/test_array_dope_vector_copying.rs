use super::helpers::run_prints;

#[test]
fn array_dope_vector_copying_alloc_to_alloc_shape_transfer() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_alloc_to_alloc_shape_transfer
    integer, allocatable :: source(:), target(:)
    source = (/ 4, 8, 12, 16 /)
    target = source
    print *, size(target)
    print *, sum(target)
    print *, target(1)
    print *, target(size(target))
end program array_dope_vector_copying_alloc_to_alloc_shape_transfer
"#,
    );
    assert_eq!(out, vec!["4", "40", "4", "16"]);
}

#[test]
fn array_dope_vector_copying_alloc_bounds_shifted_target() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_alloc_bounds_shifted_target
    integer, allocatable :: source(:)
    integer :: dest(-2:2)
    integer :: i
    source = (/ 1, 2, 3, 4, 5 /)
    dest(-2:2) = source
    print *, lbound(dest, 1)
    print *, ubound(dest, 1)
    print *, dest(-2)
    print *, dest(2)
end program array_dope_vector_copying_alloc_bounds_shifted_target
"#,
    );
    assert_eq!(out, vec!["-2", "2", "1", "5"]);
}

#[test]
fn array_dope_vector_copying_assign_from_section_to_alloc() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_assign_from_section_to_alloc
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 10, 20, 30, 40, 50, 60 /)
    target = source(2:5)
    print *, size(target)
    print *, sum(target)
    print *, target(1)
    print *, target(size(target))
end program array_dope_vector_copying_assign_from_section_to_alloc
"#,
    );
    assert_eq!(out, vec!["4", "140", "20", "50"]);
}

#[test]
fn array_dope_vector_copying_section_to_mismatched_lbound() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_section_to_mismatched_lbound
    integer, allocatable :: source(:)
    integer :: target(10:12)
    source = (/ 7, 8, 9 /)
    target = source(2:4)
    print *, source(1)
    print *, source(3)
    print *, target(10)
    print *, target(12)
end program array_dope_vector_copying_section_to_mismatched_lbound
"#,
    );
    assert_eq!(out, vec!["7", "9", "8", "9"]);
}

#[test]
fn array_dope_vector_copying_copy_with_explicit_shape_declaration() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_with_explicit_shape_declaration
    integer :: source(1:3)
    integer :: target(4:6)
    source = (/ 3, 6, 9 /)
    target = source
    print *, lbound(source, 1)
    print *, ubound(target, 1)
    print *, target(4)
    print *, target(6)
end program array_dope_vector_copying_copy_with_explicit_shape_declaration
"#,
    );
    assert_eq!(out, vec!["1", "6", "3", "9"]);
}

#[test]
fn array_dope_vector_copying_copy_between_sections_same_extent() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_between_sections_same_extent
    integer :: source(1:8)
    integer :: work(0:7)
    source = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    work(2:6) = source(3:7)
    print *, sum(work)
    print *, work(2)
    print *, work(6)
end program array_dope_vector_copying_copy_between_sections_same_extent
"#,
    );
    assert_eq!(out, vec!["22", "3", "7"]);
}

#[test]
fn array_dope_vector_copying_copy_via_temporary_expression() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_via_temporary_expression
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 5, 10, 15, 20 /)
    target = source + 1
    print *, size(target)
    print *, sum(target)
    print *, target(1)
    print *, target(size(target))
end program array_dope_vector_copying_copy_via_temporary_expression
"#,
    );
    assert_eq!(out, vec!["4", "54", "6", "21"]);
}

#[test]
fn array_dope_vector_copying_copy_of_repeated_pattern() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_of_repeated_pattern
    integer :: source(1:4)
    integer :: target(1:4)
    source = (/ 2 * 1, 2 * 4 /)
    target = source
    print *, merge(1, 0, all(target == source))
    print *, target(1) + target(4)
end program array_dope_vector_copying_copy_of_repeated_pattern
"#,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn array_dope_vector_copying_copy_before_and_after_realloc() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_before_and_after_realloc
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 1, 2, 3 /)
    target = source
    source = (/ 8, 9, 10, 11 /)
    print *, sum(target)
    print *, size(source)
    print *, size(target)
    print *, target(2)
end program array_dope_vector_copying_copy_before_and_after_realloc
"#,
    );
    assert_eq!(out, vec!["6", "4", "3", "2"]);
}

#[test]
fn array_dope_vector_copying_copy_from_assumed_shape_subroutine() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_from_assumed_shape_subroutine
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 4, 3, 2, 1 /)
    call copy_out(source, target)
    print *, size(target)
    print *, sum(target)
    print *, target(1)
    print *, target(size(target))
contains
    subroutine copy_out(values, out_values)
        integer, intent(in) :: values(:)
        integer, allocatable, intent(out) :: out_values(:)
        out_values = values
    end subroutine copy_out
end program array_dope_vector_copying_copy_from_assumed_shape_subroutine
"#,
    );
    assert_eq!(out, vec!["4", "10", "4", "1"]);
}

#[test]
fn array_dope_vector_copying_copy_into_zeroed_section_then_fill() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_into_zeroed_section_then_fill
    integer :: buffer(0:9)
    integer :: donor(2:6)
    donor = (/ 11, 22, 33, 44, 55 /)
    buffer(0:4) = 0
    buffer(2:6) = donor
    print *, sum(buffer)
    print *, buffer(0)
    print *, buffer(6)
    print *, buffer(2)
end program array_dope_vector_copying_copy_into_zeroed_section_then_fill
"#,
    );
    assert_eq!(out, vec!["165", "0", "55", "11"]);
}

#[test]
fn array_dope_vector_copying_temporary_to_alloc_assign_after_bound_calculation() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_temporary_to_alloc_assign_after_bound_calculation
    integer, allocatable :: src(:)
    integer, allocatable :: dst(:)
    integer :: n
    n = 3
    src = (/ 6, 7, 8 /)
    dst = src * n
    print *, size(dst)
    print *, sum(dst)
    print *, dst(1)
    print *, dst(3)
end program array_dope_vector_copying_temporary_to_alloc_assign_after_bound_calculation
"#,
    );
    assert_eq!(out, vec!["3", "63", "18", "24"]);
}

#[test]
fn array_dope_vector_copying_2d_shape_transfer_from_reshape() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_2d_shape_transfer_from_reshape
    integer, allocatable :: source(:)
    integer, allocatable :: target(:,:)
    integer :: t
    source = (/ 1, 2, 3, 4, 5, 6 /)
    target = reshape(source, (/2,3/))
    t = target(2, 3)
    print *, size(target,1)
    print *, size(target,2)
    print *, sum(target)
    print *, t
end program array_dope_vector_copying_2d_shape_transfer_from_reshape
"#,
    );
    assert_eq!(out, vec!["2", "3", "21", "6"]);
}

#[test]
fn array_dope_vector_copying_2d_to_1d_slice_assignment() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_2d_to_1d_slice_assignment
    integer :: matrix(2,3)
    integer :: column(2)
    matrix = reshape((/ 1, 2, 3, 4, 5, 6 /), (/2,3/))
    column = matrix(:,2)
    print *, size(column)
    print *, sum(column)
    print *, column(1)
    print *, column(2)
end program array_dope_vector_copying_2d_to_1d_slice_assignment
"#,
    );
    assert_eq!(out, vec!["2", "7", "2", "4"]);
}

#[test]
fn array_dope_vector_copying_slice_update_with_expression_copy() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_slice_update_with_expression_copy
    integer :: source(1:5)
    integer :: target(1:5)
    source = (/ 2, 4, 6, 8, 10 /)
    target = source
    target(2:4) = target(1:3) + 1
    print *, target(1)
    print *, target(2)
    print *, target(3)
    print *, target(4)
    print *, target(5)
end program array_dope_vector_copying_slice_update_with_expression_copy
"#,
    );
    assert_eq!(out, vec!["2", "5", "8", "10", "10"]);
}

#[test]
fn array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    integer :: sum_target
    source = (/ 3, 1, 4, 1, 5, 9 /)
    target = source(2:5)
    target = source
    sum_target = sum(target)
    print *, size(target)
    print *, sum_target
    print *, target(1)
    print *, target(size(target))
end program array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds
"#,
    );
    assert_eq!(out, vec!["6", "23", "3", "9"]);
}

#[test]
fn array_dope_vector_copying_assign_to_scalar_array_component_like_shape() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_assign_to_scalar_array_component_like_shape
    integer, allocatable :: source(:)
    integer :: target(0:3)
    source = (/ 1, 2, 3, 4 /)
    target = source
    print *, lbound(target, 1)
    print *, ubound(target, 1)
    print *, target(0)
    print *, target(3)
end program array_dope_vector_copying_assign_to_scalar_array_component_like_shape
"#,
    );
    assert_eq!(out, vec!["0", "3", "1", "4"]);
}

#[test]
fn array_dope_vector_copying_copy_of_zero_filled_vector() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_of_zero_filled_vector
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 0, 0, 0 /)
    target = source
    print *, size(target)
    print *, sum(target)
    print *, target(1) + target(size(target))
end program array_dope_vector_copying_copy_of_zero_filled_vector
"#,
    );
    assert_eq!(out, vec!["3", "0", "0"]);
}

#[test]
fn array_dope_vector_copying_copy_reduces_through_sum_then_rebroadcast() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_reduces_through_sum_then_rebroadcast
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    integer :: total
    source = (/ 1, 2, 3, 4, 5, 6 /)
    total = sum(source)
    target = (/ total /)
    print *, size(target)
    print *, sum(target)
    print *, target(1)
end program array_dope_vector_copying_copy_reduces_through_sum_then_rebroadcast
"#,
    );
    assert_eq!(out, vec!["1", "21", "21"]);
}

#[test]
fn array_dope_vector_copying_copy_from_reshape_then_assign_back() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_from_reshape_then_assign_back
    integer :: matrix(3,2)
    integer, allocatable :: flat(:)
    matrix = reshape((/ 9, 8, 7, 6, 5, 4 /), (/3,2/))
    flat = reshape(matrix, (/6/))
    print *, size(flat)
    print *, sum(flat)
    print *, flat(1)
    print *, flat(6)
end program array_dope_vector_copying_copy_from_reshape_then_assign_back
"#,
    );
    assert_eq!(out, vec!["6", "39", "9", "4"]);
}

#[test]
fn array_dope_vector_copying_copy_between_arrays_with_different_lower_bounds() {
    let out = run_prints(
        r#"
program array_dope_vector_copying_copy_between_arrays_with_different_lower_bounds
    integer :: left(-1:4)
    integer :: right(10:15)
    left = (/ 1, 2, 3, 4, 5, 6 /)
    right = left
    print *, lbound(left, 1)
    print *, lbound(right, 1)
    print *, right(10)
    print *, right(15)
end program array_dope_vector_copying_copy_between_arrays_with_different_lower_bounds
"#,
    );
    assert_eq!(out, vec!["-1", "10", "1", "6"]);
}
