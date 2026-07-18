use super::helpers::run_prints;

#[test]
fn array_bounds_check_failure_paths_guarded_lower_bound_scalar() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_guarded_lower_bound_scalar
    integer :: a(5)
    integer :: index
    integer :: status
    a = (/ 10, 20, 30, 40, 50 /)
    index = 0
    if (index < 1) then
        status = 0
    else
        status = a(index)
    end if
    print *, status
end program array_bounds_check_failure_paths_guarded_lower_bound_scalar
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn array_bounds_check_failure_paths_guarded_upper_bound_scalar() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_guarded_upper_bound_scalar
    integer :: a(5)
    integer :: index
    integer :: status
    a = (/ 10, 20, 30, 40, 50 /)
    index = 10
    if (index > ubound(a, 1)) then
        status = -1
    else
        status = a(index)
    end if
    print *, status
end program array_bounds_check_failure_paths_guarded_upper_bound_scalar
"#,
    );
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn array_bounds_check_failure_paths_inside_while_with_bounds() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_inside_while_with_bounds
    integer :: a(3)
    integer :: i
    integer :: total
    a = (/ 1, 2, 3 /)
    total = 0
    i = 1
    do while (i <= 5)
        if (i >= lbound(a, 1) .and. i <= ubound(a, 1)) then
            total = total + a(i)
        end if
        i = i + 1
    end do
    print *, total
end program array_bounds_check_failure_paths_inside_while_with_bounds
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn array_bounds_check_failure_paths_negative_indices_never_read() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_negative_indices_never_read
    integer :: values(-1:2)
    integer :: result
    integer :: idx
    values(-1) = 9
    values(0) = 8
    values(1) = 7
    values(2) = 6

    result = 0
    idx = -2
    if (idx >= lbound(values, 1) .and. idx <= ubound(values, 1)) then
        result = values(idx)
    else
        result = -1
    end if
    print *, result

    idx = 3
    if (idx >= lbound(values, 1) .and. idx <= ubound(values, 1)) then
        result = values(idx)
    else
        result = -2
    end if
    print *, result
end program array_bounds_check_failure_paths_negative_indices_never_read
"#,
    );
    assert_eq!(out, vec!["-1", "-2"]);
}

#[test]
fn array_bounds_check_failure_paths_assumed_shape_argument_guarded() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_assumed_shape_argument_guarded
    integer :: source(1:4)
    integer :: value
    source = (/ 4, 3, 2, 1 /)
    call read_with_guard(source, 1, value)
    print *, value
    call read_with_guard(source, 8, value)
    print *, value

contains
    subroutine read_with_guard(items, idx, value)
        integer, intent(in) :: items(:)
        integer, intent(in) :: idx
        integer, intent(out) :: value
        if (idx < lbound(items, 1) .or. idx > ubound(items, 1)) then
            value = -1
        else
            value = items(idx)
        end if
    end subroutine read_with_guard
end program array_bounds_check_failure_paths_assumed_shape_argument_guarded
"#,
    );
    assert_eq!(out, vec!["4", "-1"]);
}

#[test]
fn array_bounds_check_failure_paths_multi_dim_guarded_access() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_multi_dim_guarded_access
    integer :: grid(2,3)
    integer :: status
    integer :: i
    integer :: j
    integer :: sum

    grid = reshape((/ 1, 2, 3, 4, 5, 6 /), (/2, 3/))
    sum = 0
    i = 3
    j = 1
    if (i >= lbound(grid, 1) .and. i <= ubound(grid, 1) &
        .and. j >= lbound(grid, 2) .and. j <= ubound(grid, 2)) then
        status = grid(i, j)
    else
        status = -1
    end if
    sum = sum + status

    i = 1
    j = 5
    if (i >= lbound(grid, 1) .and. i <= ubound(grid, 1) &
        .and. j >= lbound(grid, 2) .and. j <= ubound(grid, 2)) then
        status = grid(i, j)
    else
        status = -2
    end if
    sum = sum + status

    print *, sum
end program array_bounds_check_failure_paths_multi_dim_guarded_access
"#,
    );
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn array_bounds_check_failure_paths_2d_guarded_default() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_2d_guarded_default
    integer :: matrix(4,4)
    integer :: i, j, hits

    matrix = reshape((/ (i, i = 1, 16) /), (/4,4/))
    hits = 0
    i = -1
    j = 2
    if (i >= 1 .and. i <= 4 .and. j >= 1 .and. j <= 4) hits = hits + 1
    if (i == 1 .and. j == 2) hits = hits + 1
    i = 2
    if (i >= 1 .and. i <= 4 .and. j >= 1 .and. j <= 4) hits = hits + 1
    print *, hits
end program array_bounds_check_failure_paths_2d_guarded_default
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn array_bounds_check_failure_paths_section_request_inside_bounds() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_section_request_inside_bounds
    integer :: a(1:10)
    integer :: b
    integer :: c
    a = (/ (i, i = 1, 10) /)

    if (1 >= lbound(a, 1) .and. 5 <= ubound(a, 1)) then
        b = a(1)
        c = a(5)
    else
        b = -1
        c = -1
    end if

    print *, b
    print *, c
end program array_bounds_check_failure_paths_section_request_inside_bounds
"#,
    );
    assert_eq!(out, vec!["1", "5"]);
}

#[test]
fn array_bounds_check_failure_paths_section_outside_reject() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_section_outside_reject
    integer :: a(1:8)
    integer :: status
    a = (/ (i, i = 1, 8) /)

    if (9 >= lbound(a, 1) .and. 12 <= ubound(a, 1)) then
        status = a(9)
    else
        status = -1
    end if
    print *, status

    if (7 >= lbound(a, 1) .and. 9 <= ubound(a, 1)) then
        status = a(7)
    else
        status = -1
    end if
    print *, status
end program array_bounds_check_failure_paths_section_outside_reject
"#,
    );
    assert_eq!(out, vec!["-1", "7"]);
}

#[test]
fn array_bounds_check_failure_paths_stride_section_upper_trim() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_stride_section_upper_trim
    integer :: a(1:10)
    integer :: i
    integer :: total
    integer :: next
    a = (/ (i, i = 1, 10) /)
    total = 0
    next = 1
    do while (next <= 12)
        if (next >= 1 .and. next <= ubound(a, 1)) then
            total = total + a(next)
        end if
        next = next + 3
    end do
    print *, total
end program array_bounds_check_failure_paths_stride_section_upper_trim
"#,
    );
    assert_eq!(out, vec!["22"]);
}

#[test]
fn array_bounds_check_failure_paths_stride_section_negative_step() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_stride_section_negative_step
    integer :: a(1:9)
    integer :: cursor
    integer :: total
    a = (/ (i, i = 1, 9) /)
    total = 0
    cursor = 9
    do while (cursor >= 0)
        if (cursor >= lbound(a, 1) .and. cursor <= ubound(a, 1)) then
            total = total + a(cursor)
        end if
        cursor = cursor - 4
    end do
    print *, total
end program array_bounds_check_failure_paths_stride_section_negative_step
"#,
    );
    assert_eq!(out, vec!["13"]);
}

#[test]
fn array_bounds_check_failure_paths_section_start_greater_than_end() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_section_start_greater_than_end
    integer :: a(1:6)
    integer :: sum

    a = (/ 1, 2, 3, 4, 5, 6 /)
    sum = 0
    if (5 >= lbound(a, 1) .and. 3 <= ubound(a, 1) .and. 5 <= 3) then
        sum = sum + a(5) + a(3)
    else
        sum = -99
    end if
    print *, sum
end program array_bounds_check_failure_paths_section_start_greater_than_end
"#,
    );
    assert_eq!(out, vec!["-99"]);
}

#[test]
fn array_bounds_check_failure_paths_zero_length_vector_guard() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_zero_length_vector_guard
    integer :: i
    integer :: status

    status = 0
    do i = -2, 2
        if (i >= 1 .and. i <= 0) then
            status = 1
        else
            status = status + 1
        end if
    end do
    print *, status
end program array_bounds_check_failure_paths_zero_length_vector_guard
"#,
    );
    assert_eq!(out, vec!["5"]);
}

#[test]
fn array_bounds_check_failure_paths_allocated_vector_guard() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_allocated_vector_guard
    integer, allocatable :: a(:)
    integer :: value

    allocate(a(2:6))
    a = (/ 4, 8, 15, 16, 23 /)
    if (lbound(a, 1) /= 2 .or. ubound(a, 1) /= 6) then
        value = -1
    else
        value = a(4)
    end if
    print *, value
    deallocate(a)
end program array_bounds_check_failure_paths_allocated_vector_guard
"#,
    );
    assert_eq!(out, vec!["15"]);
}

#[test]
fn array_bounds_check_failure_paths_allocated_vector_misaligned_request() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_allocated_vector_misaligned_request
    integer, allocatable :: a(:)
    integer :: result
    allocate(a(-1:1))
    a = (/ 7, 8, 9 /)

    if (a(0) == 8 .and. lbound(a, 1) == -1) then
        result = 1
    else
        result = 0
    end if
    print *, result

    if (ubound(a, 1) == 1 .and. lbound(a, 1) == -1) then
        result = result + 1
    else
        result = result - 1
    end if
    print *, result

    deallocate(a)
end program array_bounds_check_failure_paths_allocated_vector_misaligned_request
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn array_bounds_check_failure_paths_pointer_slice_bounds() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_pointer_slice_bounds
    integer, target :: source(0:9)
    integer, pointer :: alias(:)

    source = (/ (i, i = 0, 9) /)
    alias => source(2:8)

    if (lbound(alias, 1) >= 2 .and. ubound(alias, 1) <= 8) then
        print *, alias(lbound(alias, 1))
        print *, alias(ubound(alias, 1))
    else
        print *, -1
        print *, -1
    end if
end program array_bounds_check_failure_paths_pointer_slice_bounds
"#,
    );
    assert_eq!(out, vec!["2", "8"]);
}

#[test]
fn array_bounds_check_failure_paths_pointer_slice_outside_guard() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_pointer_slice_outside_guard
    integer, target :: source(10:14)
    integer, pointer :: alias(:)
    integer :: status

    source = (/ 1, 2, 3, 4, 5 /)
    alias => source(10:14)

    if (11 >= lbound(alias, 1) .and. 11 <= ubound(alias, 1)) then
        status = alias(11)
    else
        status = -1
    end if
    print *, status

    if (15 >= lbound(alias, 1) .and. 15 <= ubound(alias, 1)) then
        status = alias(15)
    else
        status = -2
    end if
    print *, status
end program array_bounds_check_failure_paths_pointer_slice_outside_guard
"#,
    );
    assert_eq!(out, vec!["1", "-2"]);
}

#[test]
fn array_bounds_check_failure_paths_nested_calls_with_invariant_bounds() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_nested_calls_with_invariant_bounds
    integer :: values(1:7)
    integer :: found

    values = (/ 2, 4, 6, 8, 10, 12, 14 /)
    call check(values, 1, found)
    print *, found
    call check(values, 10, found)
    print *, found

contains
    subroutine check(a, idx, out)
        integer, intent(in) :: a(:)
        integer, intent(in) :: idx
        integer, intent(out) :: out
        if (idx < lbound(a, 1) .or. idx > ubound(a, 1)) then
            out = -1
        else
            out = a(idx)
        end if
    end subroutine check
end program array_bounds_check_failure_paths_nested_calls_with_invariant_bounds
"#,
    );
    assert_eq!(out, vec!["2", "-1"]);
}

#[test]
fn array_bounds_check_failure_paths_scalar_guard_without_false_positive() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_scalar_guard_without_false_positive
    integer :: a(3:7)
    integer :: idx
    integer :: count

    a = (/ 11, 22, 33, 44, 55 /)
    count = 0
    do idx = 3, 7
        if (idx >= lbound(a, 1) .and. idx <= ubound(a, 1)) then
            count = count + a(idx)
        end if
    end do
    print *, count
end program array_bounds_check_failure_paths_scalar_guard_without_false_positive
"#,
    );
    assert_eq!(out, vec!["165"]);
}

#[test]
fn array_bounds_check_failure_paths_matrix_guarded_reassignment() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_matrix_guarded_reassignment
    integer :: src(1:2, 1:2)
    integer :: dst(1:2, 1:2)
    integer :: i
    integer :: j

    src = reshape((/ 1, 2, 3, 4 /), (/2,2/))
    dst = 0

    do i = 0, 2
        do j = 0, 2
            if (i >= lbound(src, 1) .and. i <= ubound(src, 1) .and. &
                j >= lbound(src, 2) .and. j <= ubound(src, 2)) then
                dst(i, j) = src(i, j)
            else
                dst(i, j) = -1
            end if
        end do
    end do

    print *, dst(1,1)
    print *, dst(0,0)
end program array_bounds_check_failure_paths_matrix_guarded_reassignment
"#,
    );
    assert_eq!(out, vec!["1", "-1"]);
}

#[test]
fn array_bounds_check_failure_paths_string_slice_like_bounds() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_string_slice_like_bounds
    character(len=8) :: word
    integer :: idx
    integer :: status

    word = 'fortran '
    idx = 0
    if (idx >= 1 .and. idx <= len(word)) then
        status = ichar(word(idx:idx))
    else
        status = -1
    end if
    print *, status

    idx = 8
    if (idx >= 1 .and. idx <= len(word)) then
        status = ichar(word(idx:idx))
    else
        status = -1
    end if
    print *, status
end program array_bounds_check_failure_paths_string_slice_like_bounds
"#,
    );
    assert_eq!(out, vec!["-1", "32"]);
}

#[test]
fn array_bounds_check_failure_paths_pack_like_masking() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_pack_like_masking
    integer :: a(1:6)
    integer :: i
    integer :: kept

    a = (/ 3, -1, 4, -2, 5, -3 /)
    kept = 0
    do i = 1, 6
        if (i >= lbound(a, 1) .and. i <= ubound(a, 1) .and. a(i) > 0) then
            kept = kept + 1
        end if
    end do
    print *, kept
end program array_bounds_check_failure_paths_pack_like_masking
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn array_bounds_check_failure_paths_selective_element_copy() {
    let out = run_prints(
        r#"
program array_bounds_check_failure_paths_selective_element_copy
    integer :: src(10:12)
    integer :: dst(1:5)
    integer :: i

    src = (/ 1, 2, 3 /)

    do i = 1, 5
        if (i >= lbound(src, 1) .and. i <= ubound(src, 1)) then
            dst(i) = src(i)
        else
            dst(i) = -1
        end if
    end do

    print *, dst(1)
    print *, dst(3)
    print *, dst(5)
end program array_bounds_check_failure_paths_selective_element_copy
"#,
    );
    assert_eq!(out, vec!["-1", "-1", "-1"]);
}
