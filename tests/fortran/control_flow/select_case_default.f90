! vybe-test: fortran/control_flow/select_case_default
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
character(len=5) :: vybe_check_w(1) = [ "other" ]
    integer :: x
    x = 99
    select case (x)
        case (1)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("one") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "one", "]"
                stop 1
            end if
        case default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("other") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "other", "]"
                stop 1
            end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
