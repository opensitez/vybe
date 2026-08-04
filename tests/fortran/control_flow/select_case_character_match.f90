! vybe-test: fortran/control_flow/select_case_character_match
! origin: languages/fortran/tests/fortran/test_control_flow.rs

program test
integer :: vybe_check_i = 0
character(len=2) :: vybe_check_w(1) = [ "ok" ]
    character(len=5) :: mode
    mode = "beta"
    select case (trim(mode))
        case ("alpha", "beta")
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("ok") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "ok", "]"
                stop 1
            end if
        case ("gamma")
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if (trim("skip") /= trim(vybe_check_w(vybe_check_i))) then
                print *, "FAIL at ", vybe_check_i, " got [", "skip", "]"
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
