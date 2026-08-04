! vybe-test: fortran/select_case_complex_ranges/select_case_complex_ranges_character_like_with_ranges
! origin: languages/fortran/tests/fortran/test_select_case_complex_ranges.rs

program select_case_complex_ranges_character_like_with_ranges
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "group" ]
    character(len=1) :: c
    c = 'b'
    select case (c)
    case ('a':'d', 'f':'z')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('group') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'group', "]"
            stop 1
        end if
    case ('e')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('alone') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'alone', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_complex_ranges_character_like_with_ranges
