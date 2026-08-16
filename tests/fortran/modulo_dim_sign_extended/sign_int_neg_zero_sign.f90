! vybe-test: fortran/modulo_dim_sign_extended/sign_int_neg_zero_sign
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((sign(-7, 0)) /= 7) then
    print *, "FAIL: want [7] got [", sign(-7, 0), "]"
    stop 1
end if
end program t
