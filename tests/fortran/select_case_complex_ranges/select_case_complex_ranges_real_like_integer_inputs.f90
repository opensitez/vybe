! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_real_like_integer_inputs
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_real_like_integer_inputs
integer :: vybe_check_i = 0
character(len=10) :: vybe_check_w(1) = [ "even-small" ]
    integer :: n
    n = 2
    select case (n)
    case (:0)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('nonpos') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'nonpos', "]"
            stop 1
        end if
    case (1, 3, 5)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('odd-small') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'odd-small', "]"
            stop 1
        end if
    case (2, 4, 6)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('even-small') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'even-small', "]"
            stop 1
        end if
    case default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('none') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'none', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_real_like_integer_inputs
