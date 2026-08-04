! vybe-test: fortran/variable_declarations_extended/implicit_none_integer_real
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: i = 4
real :: r = 2.0
if ((i + nint(r)) /= 6) then
    print *, "FAIL: want [6] got [", i + nint(r), "]"
    stop 1
end if
end program t
