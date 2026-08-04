! vybe-test: fortran/array_bounds_check_failure_paths/array_bounds_check_failure_paths_selective_element_copy
! origin: languages/fortran/tests/fortran/test_array_bounds_check_failure_paths.rs

program array_bounds_check_failure_paths_selective_element_copy
integer :: vybe_check_i = 0
integer :: vybe_check_w(3) = [ -1, -1, -1 ]
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

        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((dst(1)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", dst(1), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((dst(3)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", dst(3), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 3) then
        print *, "FAIL: more than 3 line(s)"
        stop 1
    end if
    if ((dst(5)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", dst(5), "]"
        stop 1
    end if
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program array_bounds_check_failure_paths_selective_element_copy
