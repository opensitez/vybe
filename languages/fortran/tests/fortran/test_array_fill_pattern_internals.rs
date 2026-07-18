use super::helpers::run_prints;

#[test]
fn array_fill_pattern_internals_scalar_fill_allocation() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_scalar_fill_allocation
    integer, allocatable :: values(:)
    allocate(values(1:6))
    values = 7
    print *, sum(values)
    print *, minval(values)
    print *, maxval(values)
end program array_fill_pattern_internals_scalar_fill_allocation
"#,
    );
    assert_eq!(out, vec!["42", "7", "7"]);
}

#[test]
fn array_fill_pattern_internals_scalar_fill_fixed_shape() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_scalar_fill_fixed_shape
    integer :: values(4)
    values = 5
    print *, sum(values)
    print *, values(1)
    print *, values(4)
end program array_fill_pattern_internals_scalar_fill_fixed_shape
"#,
    );
    assert_eq!(out, vec!["20", "5", "5"]);
}

#[test]
fn array_fill_pattern_internals_fill_after_realloc_grows() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_after_realloc_grows
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3 /)
    values = 0
    values = (/ 4, 4 /)
    values = 3
    print *, size(values)
    print *, sum(values)
    print *, values(2)
end program array_fill_pattern_internals_fill_after_realloc_grows
"#,
    );
    assert_eq!(out, vec!["2", "6", "3"]);
}

#[test]
fn array_fill_pattern_internals_vectorized_fill_via_repeat_expression() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_vectorized_fill_via_repeat_expression
    integer, allocatable :: values(:)
    values = (/ 3 * 0 /)
    print *, size(values)
    print *, sum(values)
    values = 11
    print *, sum(values)
    print *, values(3)
end program array_fill_pattern_internals_vectorized_fill_via_repeat_expression
"#,
    );
    assert_eq!(out, vec!["3", "0", "33", "11"]);
}

#[test]
fn array_fill_pattern_internals_fill_even_odd_by_mask() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_even_odd_by_mask
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6 /)
    where (mod(values, 2) == 0)
        values = 20
    elsewhere
        values = -20
    end where
    print *, sum(values)
    print *, values(1)
    print *, values(2)
end program array_fill_pattern_internals_fill_even_odd_by_mask
"#,
    );
    assert_eq!(out, vec!["0", "-20", "20"]);
}

#[test]
fn array_fill_pattern_internals_fill_from_constructor_then_merge() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_from_constructor_then_merge
    integer, allocatable :: values(:), mask(:)
    values = (/ 1, 2, 3, 4, 5 /)
    mask = merge(1, 0, values > 2)
    values = merge(100, values, mask == 1)
    print *, sum(values)
    print *, values(2)
    print *, values(5)
end program array_fill_pattern_internals_fill_from_constructor_then_merge
"#,
    );
    assert_eq!(out, vec!["315", "2", "100"]);
}

#[test]
fn array_fill_pattern_internals_fill_section_with_scalar_then_restore() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_section_with_scalar_then_restore
    integer, allocatable :: values(:)
    values = (/ 10, 20, 30, 40, 50 /)
    values(2:4) = 0
    values(1) = values(5)
    print *, sum(values)
    print *, values(2)
    print *, values(4)
end program array_fill_pattern_internals_fill_section_with_scalar_then_restore
"#,
    );
    assert_eq!(out, vec!["110", "0", "0"]);
}

#[test]
fn array_fill_pattern_internals_fill_2d_rows_with_scalar() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_2d_rows_with_scalar
    integer :: matrix(3,3)
    matrix = 0
    matrix(2,:) = 4
    print *, sum(matrix)
    print *, matrix(2,1)
    print *, matrix(1,1)
end program array_fill_pattern_internals_fill_2d_rows_with_scalar
"#,
    );
    assert_eq!(out, vec!["12", "4", "0"]);
}

#[test]
fn array_fill_pattern_internals_fill_2d_columns_with_scalar() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_2d_columns_with_scalar
    integer :: matrix(3,3)
    matrix = 1
    matrix(:,2) = -2
    print *, sum(matrix)
    print *, matrix(1,2)
    print *, matrix(3,2)
end program array_fill_pattern_internals_fill_2d_columns_with_scalar
"#,
    );
    assert_eq!(out, vec!["7", "-2", "-2"]);
}

#[test]
fn array_fill_pattern_internals_fill_slice_by_constant_boundaries() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_slice_by_constant_boundaries
    integer :: values(1:8)
    values = 1
    values(3:6) = 9
    print *, sum(values)
    print *, lbound(values(3:6), 1)
    print *, ubound(values(3:6), 1)
    print *, values(5)
end program array_fill_pattern_internals_fill_slice_by_constant_boundaries
"#,
    );
    assert_eq!(out, vec!["38", "3", "6", "9"]);
}

#[test]
fn array_fill_pattern_internals_fill_slice_by_variable_boundaries() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_slice_by_variable_boundaries
    integer :: values(1:10)
    integer :: start_idx
    integer :: end_idx
    values = 2
    start_idx = 4
    end_idx = 8
    values(start_idx:end_idx) = -1
    print *, sum(values)
    print *, values(start_idx)
    print *, values(end_idx)
end program array_fill_pattern_internals_fill_slice_by_variable_boundaries
"#,
    );
    assert_eq!(out, vec!["8", "-1", "-1"]);
}

#[test]
fn array_fill_pattern_internals_fill_matrix_outer_border() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_matrix_outer_border
    integer :: matrix(3,4)
    integer :: perimeter
    matrix = 0
    matrix(1,:) = 9
    matrix(3,:) = 9
    matrix(2,1) = 9
    matrix(2,4) = 9
    perimeter = sum(matrix)
    print *, perimeter
    print *, matrix(2,2)
    print *, matrix(1,2)
end program array_fill_pattern_internals_fill_matrix_outer_border
"#,
    );
    assert_eq!(out, vec!["44", "0", "9"]);
}

#[test]
fn array_fill_pattern_internals_fill_cross_pattern_then_flatten() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_cross_pattern_then_flatten
    integer :: matrix(4,4)
    integer :: total
    matrix = 0
    matrix(2,:) = 3
    matrix(1,2) = 3
    matrix(3,2) = 3
    matrix(:,3) = 3
    total = sum(matrix)
    print *, total
    print *, matrix(2,2)
    print *, matrix(1,1)
end program array_fill_pattern_internals_fill_cross_pattern_then_flatten
"#,
    );
    assert_eq!(out, vec!["30", "3", "0"]);
}

#[test]
fn array_fill_pattern_internals_fill_section_with_expression() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_section_with_expression
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 5, 6, 7, 8 /)
    do i = 2, 3
        values(i) = values(i) * values(1)
    end do
    print *, sum(values)
    print *, values(2)
    print *, values(3)
end program array_fill_pattern_internals_fill_section_with_expression
"#,
    );
    assert_eq!(out, vec!["47", "25", "35"]);
}

#[test]
fn array_fill_pattern_internals_fill_with_reshape_vector_back() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_with_reshape_vector_back
    integer, allocatable :: flat(:)
    integer :: mat(2,3)
    flat = (/ 1, 2, 3, 4, 5, 6 /)
    mat = reshape(flat, (/2,3/))
    flat = mat
    print *, sum(flat)
    print *, flat(1)
    print *, flat(6)
end program array_fill_pattern_internals_fill_with_reshape_vector_back
"#,
    );
    assert_eq!(out, vec!["21", "1", "6"]);
}

#[test]
fn array_fill_pattern_internals_fill_and_reset_with_where() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_and_reset_with_where
    integer, allocatable :: values(:)
    values = (/ 9, 8, 7, 6, 5, 4 /)
    where (values > 6)
        values = 0
    end where
    print *, sum(values)
    print *, values(1)
    print *, values(2)
end program array_fill_pattern_internals_fill_and_reset_with_where
"#,
    );
    assert_eq!(out, vec!["17", "0", "8"]);
}

#[test]
fn array_fill_pattern_internals_fill_nested_sections_chain() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_nested_sections_chain
    integer :: matrix(4,4)
    matrix = 1
    matrix(2:3,2:3) = 7
    matrix(3,1) = 9
    print *, sum(matrix)
    print *, matrix(2,2)
    print *, matrix(3,3)
    print *, matrix(4,4)
end program array_fill_pattern_internals_fill_nested_sections_chain
"#,
    );
    assert_eq!(out, vec!["34", "7", "7", "1"]);
}

#[test]
fn array_fill_pattern_internals_fill_then_replace_tail() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_then_replace_tail
    integer, allocatable :: values(:)
    values = (/ 1, 1, 1, 1, 1, 1 /)
    values(1:3) = -3
    values(4:6) = values(1:3) + 5
    print *, sum(values)
    print *, values(3)
    print *, values(6)
end program array_fill_pattern_internals_fill_then_replace_tail
"#,
    );
    assert_eq!(out, vec!["8", "2", "2"]);
}

#[test]
fn array_fill_pattern_internals_fill_non_default_origin_1d() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_non_default_origin_1d
    integer :: values(-2:2)
    values = 0
    values(-2:0) = 4
    values(1:2) = -1
    print *, sum(values)
    print *, lbound(values,1)
    print *, ubound(values,1)
    print *, values(0)
end program array_fill_pattern_internals_fill_non_default_origin_1d
"#,
    );
    assert_eq!(out, vec!["11", "-2", "2", "4"]);
}

#[test]
fn array_fill_pattern_internals_fill_non_default_origin_2d() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_non_default_origin_2d
    integer :: matrix(-1:2,-2:1)
    matrix = 1
    matrix(0,0) = 5
    matrix(1:2, -2:-1) = 3
    print *, sum(matrix)
    print *, matrix(0,0)
    print *, lbound(matrix, 1)
    print *, lbound(matrix, 2)
end program array_fill_pattern_internals_fill_non_default_origin_2d
"#,
    );
    assert_eq!(out, vec!["29", "5", "-1", "-2"]);
}

#[test]
fn array_fill_pattern_internals_fill_with_conditional_offsets() {
    let out = run_prints(
        r#"
program array_fill_pattern_internals_fill_with_conditional_offsets
    integer, allocatable :: values(:)
    integer :: i
    values = (/ 1, 2, 3, 4, 5, 6, 7 /)
    do i = lbound(values,1), ubound(values,1)
        if (mod(i,2) == 0) values(i) = values(i) + 10
    end do
    print *, values(2) + values(4) + values(6)
    print *, values(1) + values(3) + values(5) + values(7)
    print *, sum(values)
end program array_fill_pattern_internals_fill_with_conditional_offsets
"#,
    );
    assert_eq!(out, vec!["42", "16", "65"]);
}

