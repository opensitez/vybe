! vybe-test: fortran/select_case_character_ranges/select_case_character_ranges_open_ended_lower_bound
! origin: languages/fortran/tests/fortran/test_select_case_character_ranges.rs

program select_case_character_ranges_open_ended_lower_bound
integer :: vybe_check_i = 0
character(len=8) :: vybe_check_w(1) = [ "low-half" ]
    character(len=1) :: c
    c = 'e'
    select case (c)
    case (:'f')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('low-half') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'low-half', "]"
            stop 1
        end if
    case ('g':)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('high-half') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'high-half', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_character_ranges_open_ended_lower_bound
