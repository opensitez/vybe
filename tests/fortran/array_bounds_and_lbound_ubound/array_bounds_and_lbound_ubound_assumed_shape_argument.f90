! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_assumed_shape_argument
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_assumed_shape_argument
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 4, 8 ]
    integer :: data(4:8)
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((query_bounds(data)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", query_bounds(data), "]"
        stop 1
    end if

contains
    subroutine query_bounds(a)
        integer, intent(in) :: a(:)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((lbound(a, 1)) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", lbound(a, 1), "]"
            stop 1
        end if
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if ((ubound(a, 1)) /= vybe_check_w(vybe_check_i)) then
            print *, "FAIL at ", vybe_check_i, " got [", ubound(a, 1), "]"
            stop 1
        end if
    end subroutine query_bounds
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program array_bounds_and_lbound_ubound_assumed_shape_argument
