! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_negative_open_ended_range_matches_only_negative
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_negative_open_ended_range_matches_only_negative
integer :: vybe_check_i = 0
character(len=11) :: vybe_check_w(2) = [ "negative", "nonnegative" ]
    integer :: n
    n = -42
    select case (n)
    case (:-1)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('negative') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'negative', "]"
            stop 1
        end if
    case (0:)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('nonnegative') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'nonnegative', "]"
            stop 1
        end if
    end select

    n = 42
    select case (n)
    case (:-1)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('negative') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'negative', "]"
            stop 1
        end if
    case (0:)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 2) then
            print *, "FAIL: more than 2 line(s)"
            stop 1
        end if
        if (trim('nonnegative') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'nonnegative', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program select_case_complex_ranges_negative_open_ended_range_matches_only_negative
