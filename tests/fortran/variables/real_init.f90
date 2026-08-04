! vybe-test: fortran/variables/real_init
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
real :: pi = 3.14159
if (abs((pi) - 3.14159) > 1.0e-6) then
    print *, "FAIL: want [3.14159] got [", pi, "]"
    stop 1
end if
end program t
