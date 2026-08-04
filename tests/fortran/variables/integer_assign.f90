! vybe-test: fortran/variables/integer_assign
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer :: x
x = 99
if ((x) /= 99) then
    print *, "FAIL: want [99] got [", x, "]"
    stop 1
end if
end program t
