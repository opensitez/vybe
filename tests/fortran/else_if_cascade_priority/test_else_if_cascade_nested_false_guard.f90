! vybe-test: fortran/else_if_cascade_priority/test_else_if_cascade_nested_false_guard
! origin: languages/fortran/tests/fortran/test_else_if_cascade_priority.rs

program test_else_if_cascade_nested_false_guard
integer :: vybe_check_i = 0
character(len=10) :: vybe_check_w(1) = [ "neg-strong" ]
    real :: v
    v = -1.5
    if (v > 0.0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("pos") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "pos", "]"
            stop 1
        end if
    else if (v < -1.0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("neg-strong") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "neg-strong", "]"
            stop 1
        end if
    else if (abs(v) < 2.0) then
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("small") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "small", "]"
            stop 1
        end if
    else
                vybe_check_i = vybe_check_i + 1
        if (vybe_check_i > 1) then
            print *, "FAIL: more than 1 line(s)"
            stop 1
        end if
        if (trim("other") /= trim(vybe_check_w(vybe_check_i))) then
            print *, "FAIL at ", vybe_check_i, " got [", "other", "]"
            stop 1
        end if
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_else_if_cascade_nested_false_guard
