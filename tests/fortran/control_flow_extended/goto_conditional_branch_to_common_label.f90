! vybe-test: fortran/control_flow_extended/goto_conditional_branch_to_common_label
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: n = 4
if (n < 5) goto 10
if (trim('high') /= "low") then
    print *, "FAIL: want [low] got [", 'high', "]"
    stop 1
end if
goto 20
10 print *, 'low'
20 continue
end program t
