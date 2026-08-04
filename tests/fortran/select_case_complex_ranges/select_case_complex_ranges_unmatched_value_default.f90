! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_unmatched_value_default
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_unmatched_value_default
integer :: vybe_check_i = 0
character(len=7) :: vybe_check_w(1) = [ "default" ]
    integer :: n
    n = 101
    select case (n)
    case (1:10)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('low') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'low', "]"
            stop 1
        end if
    case (11:20)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('mid') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'mid', "]"
            stop 1
        end if
    case default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('default') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'default', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_unmatched_value_default
