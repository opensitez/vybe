! vybe-test: fortran/variables/parameter_constant
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer, parameter :: N = 42
if ((N) /= 42) then
    print *, "FAIL: want [42] got [", N, "]"
    stop 1
end if
end program t
