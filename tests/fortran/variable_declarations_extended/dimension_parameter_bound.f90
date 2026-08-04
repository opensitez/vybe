! vybe-test: fortran/variable_declarations_extended/dimension_parameter_bound
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, parameter :: n = 3
integer, dimension(n) :: vec
vec(n) = 100
if ((vec(n)) /= 100) then
    print *, "FAIL: want [100] got [", vec(n), "]"
    stop 1
end if
end program t
