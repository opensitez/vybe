! vybe-test: fortran/control_flow_extended/arithmetic_if_positive_label_1
! origin: languages/fortran/tests/fortran/test_control_flow_extended.rs
program t
integer :: x, y
x = 5
y = 0
if (x) 10,20,30
10 y = 10
goto 40
20 y = 20
goto 40
30 y = 30
40 continue
if ((y) /= 10) then
    print *, "FAIL: want [10] got [", y, "]"
    stop 1
end if
end program t
