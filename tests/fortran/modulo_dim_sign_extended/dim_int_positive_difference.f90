! vybe-test: fortran/modulo_dim_sign_extended/dim_int_positive_difference
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((dim(10, 3)) /= 7) then
    print *, "FAIL: want [7] got [", dim(10, 3), "]"
    stop 1
end if
end program t
