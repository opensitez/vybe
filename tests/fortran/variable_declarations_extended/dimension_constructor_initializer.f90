! vybe-test: fortran/variable_declarations_extended/dimension_constructor_initializer
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer, dimension(3) :: v = (/ 10, 20, 30 /)
if ((v(1)) /= 10) then
    print *, "FAIL: want [10] got [", v(1), "]"
    stop 1
end if
if ((v(2)) /= 20) then
    print *, "FAIL: want [20] got [", v(2), "]"
    stop 1
end if
if ((v(3)) /= 30) then
    print *, "FAIL: want [30] got [", v(3), "]"
    stop 1
end if
end program t
