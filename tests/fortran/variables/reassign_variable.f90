! vybe-test: fortran/variables/reassign_variable
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
integer :: x
x = 10
x = 20
if ((x) /= 20) then
    print *, "FAIL: want [20] got [", x, "]"
    stop 1
end if
end program t
