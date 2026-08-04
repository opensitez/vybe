! vybe-test: fortran/variable_declarations_extended/dimension_non_unit_lower_bound
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, dimension(-2:2) :: v
v(-2) = -2
v(-1) = -1
v(0) = 0
v(1) = 1
v(2) = 2
if ((v(-2)) /= -2) then
    print *, "FAIL: want [-2] got [", v(-2), "]"
    stop 1
end if
if ((v(2)) /= 2) then
    print *, "FAIL: want [2] got [", v(2), "]"
    stop 1
end if
end program t
