! vybe-test: fortran/basics/parameter_constant
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    integer, parameter :: MAX_SIZE = 100
    if ((MAX_SIZE) /= 100) then
    print *, "FAIL: want [100] got [", MAX_SIZE, "]"
    stop 1
end if
end program test
