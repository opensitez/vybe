! vybe-test: fortran/modulo_dim_sign_extended/real_modulo_pos_pos_nint15
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(modulo(7.5, 2.0)*10)) /= 15) then
    print *, "FAIL: want [15] got [", nint(modulo(7.5, 2.0)*10), "]"
    stop 1
end if
end program t
