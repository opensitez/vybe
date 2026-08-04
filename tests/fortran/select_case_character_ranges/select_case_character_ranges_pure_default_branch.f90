! vybe-test: fortran/select_case_character_ranges/select_case_character_ranges_pure_default_branch
! origin: languages/fortran/tests/fortran/test_select_case_character_ranges.rs

program select_case_character_ranges_pure_default_branch
integer :: vybe_check_i = 0
character(len=7) :: vybe_check_w(1) = [ "default" ]
    character(len=3) :: c
    c = 'xyz'
    select case (c)
    case ('a':'f')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('small') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'small', "]"
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
end program select_case_character_ranges_pure_default_branch
