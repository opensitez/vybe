! vybe-test: fortran/modulo_dim_sign_extended/sign_real_scaled_pos271
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(sign(2.71, 1.0)*100)) /= 271) then
    print *, "FAIL: want [271] got [", nint(sign(2.71, 1.0)*100), "]"
    stop 1
end if
end program t
