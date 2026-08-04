! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_multiple_values_and_range_mix
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_multiple_values_and_range_mix
integer :: vybe_check_i = 0
character(len=7) :: vybe_check_w(1) = [ "odd-set" ]
    integer :: n
    n = 13
    select case (n)
    case (1, 3, 5, 9, 13)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('odd-set') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'odd-set', "]"
            stop 1
        end if
    case (10:20)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('middle') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'middle', "]"
            stop 1
        end if
    case default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('out') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'out', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_multiple_values_and_range_mix
