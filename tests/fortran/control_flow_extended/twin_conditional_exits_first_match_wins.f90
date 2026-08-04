! vybe-test: fortran/control_flow_extended/twin_conditional_exits_first_match_wins
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x = 5, y
if (x == 5) then
y = 1
else if (x == 5) then
y = 2
else
y = 3
end if
if ((y) /= 1) then
    print *, "FAIL: want [1] got [", y, "]"
    stop 1
end if
end program t
