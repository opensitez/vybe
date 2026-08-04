! vybe-test: fortran/control_flow_extended/goto_skip_unreachable_assignment
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x = 0
goto 10
x = 999
10 continue
if ((x) /= 0) then
    print *, "FAIL: want [0] got [", x, "]"
    stop 1
end if
end program t
