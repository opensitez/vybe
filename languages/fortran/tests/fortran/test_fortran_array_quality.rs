use super::helpers::run_prints;

#[test]
fn array_quality_full_assignment() {
    let out = run_prints(
        r#"
program array_quality_full_assignment
    integer, dimension(5) :: values
    values = (/ 1, 2, 3, 4, 5 /)
    print *, values(1) + values(5)
end program array_quality_full_assignment
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_quality_constructor_sum() {
    let out = run_prints(
        r#"
program array_quality_constructor_sum
    integer, dimension(4) :: values
    values = (/ 2, 4, 6, 8 /)
    print *, values(1) + values(2) + values(3) + values(4)
end program array_quality_constructor_sum
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn array_quality_slice_sum() {
    let out = run_prints(
        r#"
program array_quality_slice_sum
    integer, dimension(6) :: values
    values = (/ 1, 2, 3, 4, 5, 6 /)
    print *, values(2:5:2)(1) + values(2:5:2)(2)
end program array_quality_slice_sum
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_quality_step_slice_indexing() {
    let out = run_prints(
        r#"
program array_quality_step_slice_indexing
    integer, dimension(8) :: values
    integer :: i, total
    values = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    total = 0
    do i = 1, 8, 2
        total = total + values(i)
    end do
    print *, total
end program array_quality_step_slice_indexing
"#,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn array_quality_two_dimensional_layout() {
    let out = run_prints(
        r#"
program array_quality_two_dimensional_layout
    integer, dimension(2,3) :: mat
    mat = reshape((/1, 2, 3, 4, 5, 6/), (/2,3/))
    print *, mat(2,3)
end program array_quality_two_dimensional_layout
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_quality_matrix_trace() {
    let out = run_prints(
        r#"
program array_quality_matrix_trace
    integer, dimension(3,3) :: m
    integer :: total
    m = reshape((/1, 0, 2, 0, 2, 0, 3, 0, 3/), (/3,3/))
    total = m(1,1) + m(2,2) + m(3,3)
    print *, total
end program array_quality_matrix_trace
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_quality_assign_via_implicit_shape() {
    let out = run_prints(
        r#"
program array_quality_assign_via_implicit_shape
    integer, dimension(:), allocatable :: values
    allocate(values(4))
    values = (/ 7, 8, 9, 10 /)
    print *, values(3)
    deallocate(values)
end program array_quality_assign_via_implicit_shape
"#,
    );
    assert_eq!(out, vec!["9"]);
}

#[test]
fn array_quality_where_mask() {
    let out = run_prints(
        r#"
program array_quality_where_mask
    integer, dimension(6) :: source
    integer, dimension(6) :: target
    source = (/ 1, 2, 3, 4, 5, 6 /)
    where (mod(source,2) == 0)
        target = source * 2
    elsewhere
        target = source
    end where
    print *, target(2), target(3)
end program array_quality_where_mask
"#,
    );
    assert_eq!(out, vec!["4", "3"]);
}

#[test]
fn array_quality_pack_like_filter() {
    let out = run_prints(
        r#"
program array_quality_pack_like_filter
    integer, dimension(5) :: input
    integer :: i
    integer :: total
    input = (/ 1, 0, 2, 0, 3 /)
    total = 0
    do i = 1, 5
        if (input(i) > 0) total = total + input(i)
    end do
    print *, total
end program array_quality_pack_like_filter
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_quality_element_replacement() {
    let out = run_prints(
        r#"
program array_quality_element_replacement
    integer, dimension(4) :: values
    integer :: i
    values = (/ 9, 4, 1, 7 /)
    do i = 1, 4
        values(i) = values(i) + 1
    end do
    print *, values(1), values(4)
end program array_quality_element_replacement
"#,
    );
    assert_eq!(out, vec!["10", "8"]);
}

#[test]
fn array_quality_section_reverse_access() {
    let out = run_prints(
        r#"
program array_quality_section_reverse_access
    integer, dimension(5) :: values
    values = (/ 5, 4, 3, 2, 1 /)
    print *, values(1), values(5)
end program array_quality_section_reverse_access
"#,
    );
    assert_eq!(out, vec!["5", "1"]);
}

#[test]
fn array_quality_transformation_count() {
    let out = run_prints(
        r#"
program array_quality_transformation_count
    integer, dimension(6) :: values
    integer :: i
    integer :: zeros
    values = (/ 0, 1, 0, 1, 0, 1 /)
    zeros = 0
    do i = 1, 6
        if (values(i) == 0) zeros = zeros + 1
    end do
    print *, zeros
end program array_quality_transformation_count
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_quality_pointer_like_aliasing() {
    let out = run_prints(
        r#"
program array_quality_pointer_like_aliasing
    integer, target, dimension(3) :: source
    integer, pointer :: head
    source = (/ 11, 22, 33 /)
    head => source(2)
    print *, head
end program array_quality_pointer_like_aliasing
"#,
    );
    assert_eq!(out, vec!["22"]);
}
