use super::helpers::run_prints;

#[test]
fn array_bounds_and_lbound_ubound_1d_default_bounds() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_1d_default_bounds
    integer :: values(6)
    print *, lbound(values, 1)
    print *, ubound(values, 1)
end program array_bounds_and_lbound_ubound_1d_default_bounds
"#,
    );
    assert_eq!(out, vec!["1", "6"]);
}

#[test]
fn array_bounds_and_lbound_ubound_1d_non_default_zero_lower() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_1d_non_default_zero_lower
    integer :: values(0:5)
    print *, lbound(values, 1)
    print *, ubound(values, 1)
end program array_bounds_and_lbound_ubound_1d_non_default_zero_lower
"#,
    );
    assert_eq!(out, vec!["0", "5"]);
}

#[test]
fn array_bounds_and_lbound_ubound_1d_negative_lower() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_1d_negative_lower
    integer :: values(-2:3)
    print *, lbound(values, 1)
    print *, ubound(values, 1)
end program array_bounds_and_lbound_ubound_1d_negative_lower
"#,
    );
    assert_eq!(out, vec!["-2", "3"]);
}

#[test]
fn array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound
    integer, allocatable :: values(:)
    allocate(values(-4:1))
    print *, lbound(values, 1)
    print *, ubound(values, 1)
    deallocate(values)
end program array_bounds_and_lbound_ubound_1d_after_allocation_with_lower_bound
"#,
    );
    assert_eq!(out, vec!["-4", "1"]);
}

#[test]
fn array_bounds_and_lbound_ubound_2d_default_bounds() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_2d_default_bounds
    integer :: grid(3, 4)
    print *, lbound(grid, 1)
    print *, ubound(grid, 1)
    print *, lbound(grid, 2)
    print *, ubound(grid, 2)
end program array_bounds_and_lbound_ubound_2d_default_bounds
"#,
    );
    assert_eq!(out, vec!["1", "3", "1", "4"]);
}

#[test]
fn array_bounds_and_lbound_ubound_2d_offset_bounds() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_2d_offset_bounds
    integer :: grid(-2:2, 10:12)
    print *, lbound(grid, 1)
    print *, ubound(grid, 1)
    print *, lbound(grid, 2)
    print *, ubound(grid, 2)
end program array_bounds_and_lbound_ubound_2d_offset_bounds
"#,
    );
    assert_eq!(out, vec!["-2", "2", "10", "12"]);
}

#[test]
fn array_bounds_and_lbound_ubound_3d_named_bounds() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_3d_named_bounds
    integer :: cube(-1:1, 2:4, 5:6)
    print *, lbound(cube, 1)
    print *, ubound(cube, 1)
    print *, lbound(cube, 2)
    print *, ubound(cube, 2)
    print *, lbound(cube, 3)
    print *, ubound(cube, 3)
end program array_bounds_and_lbound_ubound_3d_named_bounds
"#,
    );
    assert_eq!(out, vec!["-1", "1", "2", "4", "5", "6"]);
}

#[test]
fn array_bounds_and_lbound_ubound_section_from_non_default_base() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_section_from_non_default_base
    integer :: values(-5:5)
    print *, lbound(values( -2:3), 1)
    print *, ubound(values( -2:3), 1)
end program array_bounds_and_lbound_ubound_section_from_non_default_base
"#,
    );
    assert_eq!(out, vec!["-2", "3"]);
}

#[test]
fn array_bounds_and_lbound_ubound_stride_section() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_stride_section
    integer :: values(1:9)
    print *, lbound(values(1:9:2), 1)
    print *, ubound(values(1:9:2), 1)
end program array_bounds_and_lbound_ubound_stride_section
"#,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn array_bounds_and_lbound_ubound_section_on_default_base() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_section_on_default_base
    integer :: values(10)
    print *, lbound(values(3:7), 1)
    print *, ubound(values(3:7), 1)
end program array_bounds_and_lbound_ubound_section_on_default_base
"#,
    );
    assert_eq!(out, vec!["3", "7"]);
}

#[test]
fn array_bounds_and_lbound_ubound_assumed_shape_argument() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_assumed_shape_argument
    integer :: data(4:8)
    print *, query_bounds(data)

contains
    subroutine query_bounds(a)
        integer, intent(in) :: a(:)
        print *, lbound(a, 1)
        print *, ubound(a, 1)
    end subroutine query_bounds
end program array_bounds_and_lbound_ubound_assumed_shape_argument
"#,
    );
    assert_eq!(out, vec!["4", "8"]);
}

#[test]
fn array_bounds_and_lbound_ubound_multi_rank_argument() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_multi_rank_argument
    integer :: data(-2:1, 6:9, 0:0)
    call dump_bounds(data)

contains
    subroutine dump_bounds(x)
        integer, intent(in) :: x(:, :, :)
        print *, lbound(x, 1)
        print *, ubound(x, 1)
        print *, lbound(x, 2)
        print *, ubound(x, 2)
        print *, lbound(x, 3)
        print *, ubound(x, 3)
    end subroutine dump_bounds
end program array_bounds_and_lbound_ubound_multi_rank_argument
"#,
    );
    assert_eq!(out, vec!["-2", "1", "6", "9", "0", "0"]);
}

#[test]
fn array_bounds_and_lbound_ubound_referenced_from_nested_array() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_referenced_from_nested_array
    integer :: values(-3:3)
    integer :: nested(2)
    nested(1) = -3
    nested(2) = 3
    print *, lbound(values(nested(1):nested(2)), 1)
    print *, ubound(values(nested(1):nested(2)), 1)
end program array_bounds_and_lbound_ubound_referenced_from_nested_array
"#,
    );
    assert_eq!(out, vec!["-3", "3"]);
}

#[test]
fn array_bounds_and_lbound_ubound_allocatable_slice_after_assign() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_allocatable_slice_after_assign
    integer, allocatable :: buffer(:)
    integer, allocatable :: slice(:)
    allocate(buffer(1:12))
    slice => buffer(4:9)
    print *, lbound(slice, 1)
    print *, ubound(slice, 1)
    deallocate(buffer)
end program array_bounds_and_lbound_ubound_allocatable_slice_after_assign
"#,
    );
    assert_eq!(out, vec!["4", "9"]);
}

#[test]
fn array_bounds_and_lbound_ubound_pointer_slice() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_pointer_slice
    integer, target :: source(-2:6)
    integer, pointer :: alias(:)
    alias => source(-2:6)
    print *, lbound(alias, 1)
    print *, ubound(alias, 1)
end program array_bounds_and_lbound_ubound_pointer_slice
"#,
    );
    assert_eq!(out, vec!["-2", "6"]);
}

#[test]
fn array_bounds_and_lbound_ubound_zero_size_array() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_zero_size_array
    integer :: values(0:0)
    print *, lbound(values, 1)
    print *, ubound(values, 1)
end program array_bounds_and_lbound_ubound_zero_size_array
"#,
    );
    assert_eq!(out, vec!["0", "0"]);
}

#[test]
fn array_bounds_and_lbound_ubound_zero_to_negative_extent() {
    let out = run_prints(
        r#"
program array_bounds_and_lbound_ubound_zero_to_negative_extent
    integer :: values(-2:-2)
    print *, lbound(values, 1)
    print *, ubound(values, 1)
end program array_bounds_and_lbound_ubound_zero_to_negative_extent
"#,
    );
    assert_eq!(out, vec!["-2", "-2"]);
}
