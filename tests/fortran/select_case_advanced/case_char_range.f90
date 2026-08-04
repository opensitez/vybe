! vybe-test: fortran/select_case_advanced/case_char_range
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
integer :: vybe_check_i = 0
character(len=10) :: vybe_check_w(1) = [ "first half" ]
    character :: c = 'm'
    select case (c)
    case ('a':'m')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('first half') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'first half', "]"
            stop 1
        end if
    case ('n':'z')
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim('second half') /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", 'second half', "]"
            stop 1
        end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
