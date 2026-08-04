! vybe-test: fortran/control/ctrl_goto_computed_10
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: k, x
k = 2
x = 99
go to (10,20,30), k
10 x = 10
goto 40
20 x = 20
goto 40
30 x = 30
40 continue
if ((x) /= 20) then
    print *, "FAIL: want [20] got [", x, "]"
    stop 1
end if
end program p
