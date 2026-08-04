! vybe-test: fortran/select_case_advanced/case_advanced_character_open_ended_ranges
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
integer :: vybe_check_i = 0
character(len=6) :: vybe_check_w(1) = [ "prefix" ]
    character(len=1) :: c = 'c'
    select case (c)
    case (:'m')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('prefix') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'prefix', "]"
            stop 1
        end if
    case ('n':)
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('suffix') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'suffix', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
