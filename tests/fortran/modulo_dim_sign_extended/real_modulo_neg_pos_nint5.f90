! vybe-test: fortran/modulo_dim_sign_extended/real_modulo_neg_pos_nint5
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(modulo(-7.5, 2.0)*10)) /= 5) then
    print *, "FAIL: want [5] got [", nint(modulo(-7.5, 2.0)*10), "]"
    stop 1
end if
end program t
