! vybe-test: fortran/variables/real_assign
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
real :: x
x = 2.5
if (abs((x) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", x, "]"
    stop 1
end if
end program t
