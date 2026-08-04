! vybe-test: fortran/select_case_character_ranges/select_case_character_ranges_ascii_boundary
! origin: languages/fortran/tests/fortran/test_select_case_character_ranges.rs

program select_case_character_ranges_ascii_boundary
integer :: vybe_check_i = 0
character(len=1) :: vybe_check_w(1) = [ "Z" ]
    character(len=1) :: c
    c = 'z'
    select case (c)
    case ('a':'m')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('A') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'A', "]"
            stop 1
        end if
    case ('n':'z')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('Z') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'Z', "]"
            stop 1
        end if
    case default
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('D') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'D', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_character_ranges_ascii_boundary
