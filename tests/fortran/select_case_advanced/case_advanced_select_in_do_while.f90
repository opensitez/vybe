! vybe-test: fortran/select_case_advanced/case_advanced_select_in_do_while
! origin: languages/fortran/tests/fortran/test_select_case_advanced.rs

program test
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(3) = [ "one", "two", "other" ]
    integer :: i
    i = 1
    do while (i <= 3)
        select case (i)
        case (1)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if (trim('one') /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", 'one', "]"
                stop 1
            end if
        case (2)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if (trim('two') /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", 'two', "]"
                stop 1
            end if
        case default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 3) then
                print *, "FAIL: more than 3 line(s)"
                stop 1
            end if
            if (trim('other') /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", 'other', "]"
                stop 1
            end if
        end select
        i = i + 1
    end do
if (vybe_check_i /= 3) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 3"
    stop 1
end if
end program test
