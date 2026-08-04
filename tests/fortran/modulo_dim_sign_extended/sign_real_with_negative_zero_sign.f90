! vybe-test: fortran/modulo_dim_sign_extended/sign_real_with_negative_zero_sign
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(sign(9.0, -0.0)*10)) /= -90) then
    print *, "FAIL: want [-90] got [", nint(sign(9.0, -0.0)*10), "]"
    stop 1
end if
end program t
