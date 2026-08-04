! vybe-test: fortran/control_flow/select_case_open_ranges
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
character(len=3) :: vybe_check_w(1) = [ "low" ]
    integer :: x
    x = 2
    select case (x)
        case (:4)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("low") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "low", "]"
                stop 1
            end if
        case (5:)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("high") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "high", "]"
                stop 1
            end if
        case default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("mid") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "mid", "]"
                stop 1
            end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
