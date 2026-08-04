! vybe-test: fortran/modulo_dim_sign_extended/dim_int_both_negative
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((dim(-8, -3)) /= 0) then
    print *, "FAIL: want [0] got [", dim(-8, -3), "]"
    stop 1
end if
end program t
