! vybe-test: fortran/modulo_dim_sign_extended/sign_int_neg_to_pos
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((sign(-5, 1)) /= 5) then
    print *, "FAIL: want [5] got [", sign(-5, 1), "]"
    stop 1
end if
end program t
