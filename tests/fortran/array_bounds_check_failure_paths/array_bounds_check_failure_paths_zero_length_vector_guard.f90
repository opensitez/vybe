! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_zero_length_vector_guard
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_zero_length_vector_guard
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 5 ]
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
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((status) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", status, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program array_bounds_check_failure_paths_zero_length_vector_guard
