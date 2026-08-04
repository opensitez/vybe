! vybe-test: fortran/modulo_dim_sign_extended/dim_real_equal_scaled
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(dim(4.0, 4.0)*10)) /= 0) then
    print *, "FAIL: want [0] got [", nint(dim(4.0, 4.0)*10), "]"
    stop 1
end if
end program t
