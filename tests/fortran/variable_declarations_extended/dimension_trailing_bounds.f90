! vybe-test: fortran/variable_declarations_extended/dimension_trailing_bounds
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: grid(2, 2)
grid(1, 2) = 15
if ((grid(1, 2)) /= 15) then
    print *, "FAIL: want [15] got [", grid(1, 2), "]"
    stop 1
end if
end program t
