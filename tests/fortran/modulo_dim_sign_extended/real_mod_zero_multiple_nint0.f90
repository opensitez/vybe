! vybe-test: fortran/modulo_dim_sign_extended/real_mod_zero_multiple_nint0
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(mod(6.0, 3.0)*10)) /= 0) then
    print *, "FAIL: want [0] got [", nint(mod(6.0, 3.0)*10), "]"
    stop 1
end if
end program t
