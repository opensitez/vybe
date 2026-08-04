! vybe-test: fortran/modulo_dim_sign_extended/dim_real_scaled_nint73
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(dim(10.5, 3.2)*10)) /= 73) then
    print *, "FAIL: want [73] got [", nint(dim(10.5, 3.2)*10), "]"
    stop 1
end if
end program t
