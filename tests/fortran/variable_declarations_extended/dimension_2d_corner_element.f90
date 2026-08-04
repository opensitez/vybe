! vybe-test: fortran/variable_declarations_extended/dimension_2d_corner_element
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, dimension(2, 3) :: m
m(2, 3) = 42
if ((m(2, 3)) /= 42) then
    print *, "FAIL: want [42] got [", m(2, 3), "]"
    stop 1
end if
end program t
