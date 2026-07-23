use super::helpers::run_prints;

#[test]
fn array_section_shape_and_strides_rank_one_unit_stride() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_rank_one_unit_stride
    integer :: values(10)
    values = (/ (i, i = 1, 10) /)
    print *, lbound(values(2:8),1)
    print *, ubound(values(2:8),1)
    print *, size(values(2:8))
    print *, sum(values(2:8))
end program array_section_shape_and_strides_rank_one_unit_stride
"#,
    );
    assert_eq!(out, vec!["2", "8", "7", "35"]);
}

#[test]
fn array_section_shape_and_strides_step_two_stride() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_step_two_stride
    integer :: values(1:12)
    values = (/ (i, i = 1, 12) /)
    print *, lbound(values(2:10:2),1)
    print *, size(values(2:10:2))
    print *, values(2:10:2)(1)
    print *, values(2:10:2)(size(values(2:10:2)))
    print *, sum(values(2:10:2))
end program array_section_shape_and_strides_step_two_stride
"#,
    );
    assert_eq!(out, vec!["2", "5", "2", "10", "30"]);
}

#[test]
fn array_section_shape_and_strides_negative_stride() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_negative_stride
    integer :: values(1:9)
    values = (/ (i, i = 1, 9) /)
    print *, lbound(values(9:1:-2),1)
    print *, ubound(values(9:1:-2),1)
    print *, size(values(9:1:-2))
    print *, sum(values(9:1:-2))
end program array_section_shape_and_strides_negative_stride
"#,
    );
    assert_eq!(out, vec!["9", "1", "5", "25"]);
}

#[test]
fn array_section_shape_and_strides_two_dimensional_section_shape() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_two_dimensional_section_shape
    integer :: matrix(4,5)
    matrix = reshape((/ (i, i = 1, 20) /), (/4,5/))
    print *, lbound(matrix(2:3,1:3),1)
    print *, ubound(matrix(2:3,1:3),1)
    print *, lbound(matrix(2:3,1:3),2)
    print *, ubound(matrix(2:3,1:3),2)
    print *, sum(matrix(2:3,1:3))
end program array_section_shape_and_strides_two_dimensional_section_shape
"#,
    );
    assert_eq!(out, vec!["2", "3", "1", "3", "54"]);
}

#[test]
fn array_section_shape_and_strides_section_of_section() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_section_of_section
    integer :: matrix(5,5)
    integer :: subtotal
    matrix = reshape((/ (i, i = 1, 25) /), (/5,5/))
    subtotal = sum(matrix(4:2:-1, 3:5))
    print *, subtotal
    print *, lbound(matrix(4:2:-1,3:5),1)
    print *, ubound(matrix(4:2:-1,3:5),1)
end program array_section_shape_and_strides_section_of_section
"#,
    );
    assert_eq!(out, vec!["72", "4", "2"]);
}

#[test]
fn array_section_shape_and_strides_strided_column_slice() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_strided_column_slice
    integer :: matrix(3,6)
    matrix = reshape((/ (i, i = 1, 18) /), (/3,6/))
    print *, size(matrix(:,2:6:2),1)
    print *, size(matrix(:,2:6:2),2)
    print *, sum(matrix(:,2:6:2))
    print *, matrix(2,2:6:2)(1)
    print *, matrix(2,2:6:2)(size(matrix(:,2:6:2),2))
end program array_section_shape_and_strides_strided_column_slice
"#,
    );
    assert_eq!(out, vec!["3", "3", "60", "8", "14"]);
}

#[test]
fn array_section_shape_and_strides_vector_row_block() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_vector_row_block
    integer :: matrix(4,4)
    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    print *, size(matrix(2:3,2:4),1)
    print *, size(matrix(2:3,2:4),2)
    print *, matrix(2,2:4)(2)
    print *, sum(matrix(2:3,2:4))
end program array_section_shape_and_strides_vector_row_block
"#,
    );
    assert_eq!(out, vec!["2", "3", "7", "58"]);
}

#[test]
fn array_section_shape_and_strides_section_offset_bounds() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_section_offset_bounds
    integer :: values(-3:3)
    values = (/ (i, i = -3, 3) /)
    print *, lbound(values(-1:3),1)
    print *, ubound(values(-1:3),1)
    print *, size(values(-1:3))
    print *, sum(values(-1:3))
end program array_section_shape_and_strides_section_offset_bounds
"#,
    );
    assert_eq!(out, vec!["-1", "3", "5", "10"]);
}

#[test]
fn array_section_shape_and_strides_matrix_triangle_view() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_matrix_triangle_view
    integer :: matrix(4,4)
    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    print *, size(matrix(1:4:1,1:4:1),1)
    print *, sum(matrix(2:4:1,2:4:1))
    print *, matrix(2:4:1,2:4:1)(2,2)
    print *, matrix(2:4:1,2:4:1)(1,3)
end program array_section_shape_and_strides_matrix_triangle_view
"#,
    );
    assert_eq!(out, vec!["3", "85", "10", "12"]);
}

#[test]
fn array_section_shape_and_strides_section_shape_reassigned() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_section_shape_reassigned
    integer :: source(1:12)
    integer :: target(1:4)
    source = (/ (i, i = 1, 12) /)
    target = source(2:8:2)
    print *, size(target)
    print *, sum(target)
    print *, target(1)
    print *, target(4)
end program array_section_shape_and_strides_section_shape_reassigned
"#,
    );
    assert_eq!(out, vec!["4", "20", "2", "8"]);
}

#[test]
fn array_section_shape_and_strides_subsection_for_sum() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_subsection_for_sum
    integer :: values(1:15)
    values = (/ (i, i = 1, 15) /)
    print *, sum(values(5:15:5))
    print *, size(values(5:15:5))
    print *, values(5:15:5)(1)
    print *, values(5:15:5)(3)
end program array_section_shape_and_strides_subsection_for_sum
"#,
    );
    assert_eq!(out, vec!["30", "3", "5", "15"]);
}

#[test]
fn array_section_shape_and_strides_nested_indexed_sections() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_nested_indexed_sections
    integer :: matrix(6,6)
    integer :: subtotal
    matrix = reshape((/ (i, i = 1, 36) /), (/6,6/))
    subtotal = sum(matrix(2:6:2, 1:5:2))
    print *, subtotal
    print *, lbound(matrix(2:6:2, 1:5:2),1)
    print *, ubound(matrix(2:6:2, 1:5:2),1)
    print *, lbound(matrix(2:6:2, 1:5:2),2)
    print *, ubound(matrix(2:6:2, 1:5:2),2)
end program array_section_shape_and_strides_nested_indexed_sections
"#,
    );
    assert_eq!(out, vec!["108", "2", "6", "1", "5"]);
}

#[test]
fn array_section_shape_and_strides_strided_assign_then_fill() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_strided_assign_then_fill
    integer :: values(1:20)
    integer :: section_sum
    values = 0
    values(2:20:3) = 4
    section_sum = sum(values(2:20:3))
    print *, section_sum
    print *, size(values(2:20:3))
    print *, values(2)
    print *, values(20)
end program array_section_shape_and_strides_strided_assign_then_fill
"#,
    );
    assert_eq!(out, vec!["28", "7", "4", "4"]);
}

#[test]
fn array_section_shape_and_strides_column_vector_projection() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_column_vector_projection
    integer :: matrix(3,4)
    integer :: projected_sum
    matrix = reshape((/ (i, i = 1, 12) /), (/3,4/))
    projected_sum = sum(matrix(:,3))
    print *, projected_sum
    print *, lbound(matrix(:,3),1)
    print *, ubound(matrix(:,3),1)
    print *, matrix(1,3) + matrix(3,3)
end program array_section_shape_and_strides_column_vector_projection
"#,
    );
    assert_eq!(out, vec!["27", "1", "3", "9"]);
}

#[test]
fn array_section_shape_and_strides_row_vector_projection() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_row_vector_projection
    integer :: matrix(4,3)
    integer :: projected_sum
    matrix = reshape((/ (i, i = 1, 12) /), (/4,3/))
    projected_sum = sum(matrix(3,:))
    print *, projected_sum
    print *, lbound(matrix(3,:),1)
    print *, ubound(matrix(3,:),1)
    print *, matrix(3,1)
    print *, matrix(3,3)
end program array_section_shape_and_strides_row_vector_projection
"#,
    );
    assert_eq!(out, vec!["25", "1", "3", "9", "11"]);
}

#[test]
fn array_section_shape_and_strides_reshaped_section_compat() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_reshaped_section_compat
    integer :: matrix(2,6)
    integer :: section_sum
    matrix = reshape((/ (i, i = 1, 12) /), (/2,6/))
    section_sum = sum(reshape(matrix(1:2,2:5), (/2,2/)))
    print *, section_sum
    print *, lbound(matrix(1:2,2:5),1)
    print *, ubound(matrix(1:2,2:5),2)
end program array_section_shape_and_strides_reshaped_section_compat
"#,
    );
    assert_eq!(out, vec!["30", "1", "4"]);
}

#[test]
fn array_section_shape_and_strides_stride_and_mask_pair() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_stride_and_mask_pair
    integer :: values(1:9)
    integer :: selected_sum
    integer :: selected_count
    values = (/ (i, i = 1, 9) /)
    selected_sum = sum(values(1:9:2), values(1:9:2) > 5)
    selected_count = count(values(1:9:2) > 5)
    print *, selected_sum
    print *, selected_count
    print *, size(values(1:9:2))
end program array_section_shape_and_strides_stride_and_mask_pair
"#,
    );
    assert_eq!(out, vec!["7", "1", "5"]);
}

#[test]
fn array_section_shape_and_strides_flattened_section_sum() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_flattened_section_sum
    integer :: matrix(3,3)
    integer :: total
    matrix = reshape((/ (i, i = 1, 9) /), (/3,3/))
    total = sum(reshape(matrix(2:3,1:3), (/6/))
    print *, total
    print *, size(reshape(matrix(2:3,1:3), (/6/)))
end program array_section_shape_and_strides_flattened_section_sum
"#,
    );
    assert_eq!(out, vec!["33", "6"]);
}

#[test]
fn array_section_shape_and_strides_matrix_section_length_probe() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_matrix_section_length_probe
    integer :: matrix(5,5)
    matrix = reshape((/ (i, i = 1, 25) /), (/5,5/))
    print *, size(matrix(2:4,3:5))
    print *, size(matrix(2:4,3:5),1)
    print *, size(matrix(2:4,3:5),2)
    print *, matrix(2,3)
    print *, matrix(4,5)
end program array_section_shape_and_strides_matrix_section_length_probe
"#,
    );
    assert_eq!(out, vec!["6", "3", "2", "3", "25"]);
}

#[test]
fn array_section_shape_and_strides_section_as_lhs_rhs_same_shape() {
    let out = run_prints(
        r#"
program array_section_shape_and_strides_section_as_lhs_rhs_same_shape
    integer :: values(1:10)
    integer :: sample(2:8)
    values = (/ (i, i = 1, 10) /)
    sample = values(2:8)
    sample(1:3) = sample(1:3) + 1
    print *, sum(sample)
    print *, sample(1)
    print *, sample(7)
end program array_section_shape_and_strides_section_as_lhs_rhs_same_shape
"#,
    );
    assert_eq!(out, vec!["39", "3", "9"]);
}
