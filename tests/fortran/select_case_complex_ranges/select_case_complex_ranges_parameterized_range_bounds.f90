! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_parameterized_range_bounds
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_parameterized_range_bounds
integer :: vybe_check_i = 0
character(len=10) :: vybe_check_w(1) = [ "parametric" ]
    integer, parameter :: lo = 10
    integer, parameter :: hi = 20
    integer :: n
    n = 15
    select case (n)
    case (lo:hi)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('parametric') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'parametric', "]"
            stop 1
        end if
    case default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('fallback') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'fallback', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_parameterized_range_bounds
