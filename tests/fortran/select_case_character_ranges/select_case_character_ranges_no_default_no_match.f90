! vybe-test: fortran/select_case_character_ranges/select_case_character_ranges_no_default_no_match
! origin: languages/fortran/tests/fortran/test_select_case_character_ranges.rs

program select_case_character_ranges_no_default_no_match
integer :: vybe_check_i = 0
character(len=4) :: vybe_check_w(1) = [ "done" ]
    character(len=2) :: c
    c = 'zZ'
    select case (c)
    case ('a')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('alpha') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'alpha', "]"
            stop 1
        end if
    case ('b')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('bravo') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'bravo', "]"
            stop 1
        end if
    end select
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if (trim('done') /= trim(vybe_check_w(vybe_check_i))) then
        print *, "FAIL at ", vybe_check_i, " got [", 'done', "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program select_case_character_ranges_no_default_no_match
