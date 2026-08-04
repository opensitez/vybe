! vybe-test: fortran/control/ctrl_associate_01
! origin: languages/fortran/tests/fortran/test_control.rs
program p
integer :: x, y
x = 1
associate (a => x)
 a = a + 4
 y = a
end associate
if ((y) /= 5) then
    print *, "FAIL: want [5] got [", y, "]"
    stop 1
end if
if ((x) /= 5) then
    print *, "FAIL: want [5] got [", x, "]"
    stop 1
end if
end program p
