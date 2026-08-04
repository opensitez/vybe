! vybe-test: fortran/variables/implicit_none_compiles
! origin: languages/fortran/tests/fortran/test_variables.rs
program t
implicit none
integer :: x = 1
if ((x) /= 1) then
    print *, "FAIL: want [1] got [", x, "]"
    stop 1
end if
end program t
