! vybe-test: fortran/modulo_dim_sign_extended/sign_real_neg_magnitude_pos_sign
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(sign(-4.5, 2.0)*10)) /= 45) then
    print *, "FAIL: want [45] got [", nint(sign(-4.5, 2.0)*10), "]"
    stop 1
end if
end program t
