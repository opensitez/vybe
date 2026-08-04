! vybe-test: fortran/modulo_dim_sign_extended/sign_real_zero_magnitude
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(sign(0.0, -9.0)*10)) /= 0) then
    print *, "FAIL: want [0] got [", nint(sign(0.0, -9.0)*10), "]"
    stop 1
end if
end program t
