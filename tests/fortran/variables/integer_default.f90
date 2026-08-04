! vybe-test: fortran/variables/integer_default
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer :: x = 0
if ((x) /= 0) then
    print *, "FAIL: want [0] got [", x, "]"
    stop 1
end if
end program t
