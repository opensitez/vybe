! vybe-test: fortran/modulo_dim_sign_extended/real_mod_vs_modulo_neg_divisor
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(mod(11.5, -4.0)*10)) /= 15) then
    print *, "FAIL: want [15] got [", nint(mod(11.5, -4.0)*10), "]"
    stop 1
end if
if ((nint(modulo(11.5, -4.0)*10)) /= -5) then
    print *, "FAIL: want [-5] got [", nint(modulo(11.5, -4.0)*10), "]"
    stop 1
end if
end program t
