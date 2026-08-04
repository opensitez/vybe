! vybe-test: fortran/variable_declarations_extended/init_integer_negative
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: x = -12
if ((x) /= -12) then
    print *, "FAIL: want [-12] got [", x, "]"
    stop 1
end if
end program t
