! vybe-test: fortran/control/ctrl_arith_if_12
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: x, y
x = 1
if (x) 10,20,30
10 y = 10
goto 40
20 y = 20
goto 40
30 y = 30
40 continue
if ((y) /= 30) then
    print *, "FAIL: want [30] got [", y, "]"
    stop 1
end if
end program p
