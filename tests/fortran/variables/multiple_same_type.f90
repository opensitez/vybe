! vybe-test: fortran/variables/multiple_same_type
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer :: a = 1, b = 2, c = 3
if ((a + b + c) /= 6) then
    print *, "FAIL: want [6] got [", a + b + c, "]"
    stop 1
end if
end program t
