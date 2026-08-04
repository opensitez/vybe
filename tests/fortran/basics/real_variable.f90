! vybe-test: fortran/basics/real_variable
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    real :: pi
    pi = 3.14159
    if (abs((pi) - 3.14159) > 1.0e-6) then
    print *, "FAIL: want [3.14159] got [", pi, "]"
    stop 1
end if
end program test
