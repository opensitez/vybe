! vybe-test: fortran/modulo_dim_sign_extended/dim_with_variables
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
integer :: x=14, y=9
if ((dim(x, y)) /= 5) then
    print *, "FAIL: want [5] got [", dim(x, y), "]"
    stop 1
end if
end program t
