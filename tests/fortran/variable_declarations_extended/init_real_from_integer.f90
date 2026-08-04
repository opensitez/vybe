! vybe-test: fortran/variable_declarations_extended/init_real_from_integer
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
real :: r = 7
if ((r) /= 7) then
    print *, "FAIL: want [7] got [", r, "]"
    stop 1
end if
end program t
