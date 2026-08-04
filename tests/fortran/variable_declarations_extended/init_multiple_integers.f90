! vybe-test: fortran/variable_declarations_extended/init_multiple_integers
! origin: languages/fortran/tests/fortran/test_variable_declarations_extended.rs
program t
implicit none
integer :: p = 10, q = 20
if ((p + q) /= 30) then
    print *, "FAIL: want [30] got [", p + q, "]"
    stop 1
end if
end program t
