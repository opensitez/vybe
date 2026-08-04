! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_no_match_without_default_keeps_other_output_only
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_no_match_without_default_keeps_other_output_only
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "after" ]
    integer :: n
    n = 99
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
    case (20:30)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('mid') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'mid', "]"
            stop 1
        end if
    end select
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if (trim('after') /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", 'after', "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_no_match_without_default_keeps_other_output_only
