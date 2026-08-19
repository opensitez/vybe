! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_stride_section_negative_step
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_stride_section_negative_step
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
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
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_bounds_check_failure_paths_stride_section_negative_step
