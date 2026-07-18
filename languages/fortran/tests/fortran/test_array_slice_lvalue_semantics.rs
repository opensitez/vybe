use super::helpers::run_prints;

#[test]
fn array_slice_lvalue_semantics_direct_assignment_updates_elements() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_direct_assignment_updates_elements
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    values(2:5) = 0
    print *, sum(values)
    print *, values(1)
    print *, values(6)
end program array_slice_lvalue_semantics_direct_assignment_updates_elements
"#,
    );
    assert_eq!(out, vec!["7", "1", "6"]);
}

#[test]
fn array_slice_lvalue_semantics_strided_section_assignment() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_strided_section_assignment
    integer :: values(1:8)
    values = (/1, 2, 3, 4, 5, 6, 7, 8/)
    values(2:8:2) = 1
    print *, values(2)
    print *, values(4)
    print *, values(6)
    print *, values(8)
    print *, sum(values)
end program array_slice_lvalue_semantics_strided_section_assignment
"#,
    );
    assert_eq!(out, vec!["1", "1", "1", "1", "20"]);
}

#[test]
fn array_slice_lvalue_semantics_matrix_subsection_increment() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_matrix_subsection_increment
    integer :: values(3, 3)
    values = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    values(2:3, 2:3) = values(2:3, 2:3) + 1
    print *, values(2, 2)
    print *, values(3, 3)
    print *, sum(values)
end program array_slice_lvalue_semantics_matrix_subsection_increment
"#,
    );
    assert_eq!(out, vec!["6", "10", "53"]);
}

#[test]
fn array_slice_lvalue_semantics_row_column_split_update() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_row_column_split_update
    integer :: values(4, 4)
    values = reshape((/ (i, i = 1, 16) /), (/4, 4/))
    values(1, 2:4) = 9
    values(2:4, 1) = 7
    print *, values(1, 2)
    print *, values(1, 4)
    print *, values(4, 1)
    print *, sum(values)
end program array_slice_lvalue_semantics_row_column_split_update
"#,
    );
    assert_eq!(out, vec!["9", "9", "7", "110"]);
}

#[test]
fn array_slice_lvalue_semantics_subsection_copy_by_shape() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_subsection_copy_by_shape
    integer :: source(1:4)
    integer :: target(2, 2)
    source = (/10, 11, 12, 13/)
    target = reshape(source, (/2, 2/))
    target(1, :) = 0
    print *, target(1, 1)
    print *, target(1, 2)
    print *, target(2, 1)
    print *, target(2, 2)
end program array_slice_lvalue_semantics_subsection_copy_by_shape
"#,
    );
    assert_eq!(out, vec!["0", "0", "12", "13"]);
}

#[test]
fn array_slice_lvalue_semantics_mixed_stride_reassign() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_mixed_stride_reassign
    integer :: values(1:10)
    values = (/1, 1, 1, 1, 1, 1, 1, 1, 1, 1/)
    values(2:10:3) = (/2, 3, 4/)
    print *, values(2)
    print *, values(5)
    print *, values(8)
    print *, sum(values)
end program array_slice_lvalue_semantics_mixed_stride_reassign
"#,
    );
    assert_eq!(out, vec!["2", "3", "4", "18"]);
}

#[test]
fn array_slice_lvalue_semantics_section_as_ghost_target() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_section_as_ghost_target
    integer :: values(1:6)
    values = (/1, 2, 3, 4, 5, 6/)
    values(4:) = 0
    print *, values(3)
    print *, values(4)
    print *, values(6)
    print *, sum(values)
end program array_slice_lvalue_semantics_section_as_ghost_target
"#,
    );
    assert_eq!(out, vec!["3", "0", "0", "9"]);
}

#[test]
fn array_slice_lvalue_semantics_2d_reused_section_in_expression() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_2d_reused_section_in_expression
    integer :: values(3, 3)
    values = reshape((/1, 2, 3, 4, 5, 6, 7, 8, 9/), (/3, 3/))
    values(2:3, :) = values(1:2, :) + 1
    print *, values(2, 1)
    print *, values(3, 3)
    print *, sum(values)
end program array_slice_lvalue_semantics_2d_reused_section_in_expression
"#,
    );
    assert_eq!(out, vec!["2", "8", "45"]);
}

#[test]
fn array_slice_lvalue_semantics_vector_alias_preserved_after_reassign() {
    let out = run_prints(
        r#"
program array_slice_lvalue_semantics_vector_alias_preserved_after_reassign
    integer :: a(1:6)
    integer :: b(1:6)
    a = (/1, 2, 3, 4, 5, 6/)
    b = a
    a(2:5:2) = b(1:2)
    print *, a(2)
    print *, a(4)
    print *, sum(a)
end program array_slice_lvalue_semantics_vector_alias_preserved_after_reassign
"#,
    );
    assert_eq!(out, vec!["1", "2", "21"]);
}
