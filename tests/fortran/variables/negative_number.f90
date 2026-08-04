! vybe-test: fortran/variables/negative_number
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer :: x = -5
if ((x) /= -5) then
    print *, "FAIL: want [-5] got [", x, "]"
    stop 1
end if
end program t
