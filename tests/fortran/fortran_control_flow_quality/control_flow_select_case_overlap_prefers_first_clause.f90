! vybe-test: fortran/fortran_control_flow_quality/control_flow_select_case_overlap_prefers_first_clause
! origin: languages/fortran/tests/fortran/test_fortran_control_flow_quality.rs

program control_flow_select_case_overlap_prefers_first_clause
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 1 ]
    integer :: n
    n = 5
    select case (n)
        case (1:10)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((1) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 1, "]"
                stop 1
            end if
        case (5)
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((2) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 2, "]"
                stop 1
            end if
        case default
                        vybe_check_i = vybe_check_i + 1
            if (vybe_check_i > 1) then
                print *, "FAIL: more than 1 line(s)"
                stop 1
            end if
            if ((3) /= vybe_check_w(vybe_check_i)) then
                print *, "FAIL at ", vybe_check_i, " got [", 3, "]"
                stop 1
            end if
    end select
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program control_flow_select_case_overlap_prefers_first_clause
