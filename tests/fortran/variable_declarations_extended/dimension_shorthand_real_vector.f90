! vybe-test: fortran/variable_declarations_extended/dimension_shorthand_real_vector
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
real :: v(5)
v(5) = 9.0
if ((v(5)) /= 9) then
    print *, "FAIL: want [9] got [", v(5), "]"
    stop 1
end if
end program t
