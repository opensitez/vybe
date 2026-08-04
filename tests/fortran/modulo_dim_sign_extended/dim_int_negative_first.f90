! vybe-test: fortran/modulo_dim_sign_extended/dim_int_negative_first
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((dim(-2, 5)) /= 0) then
    print *, "FAIL: want [0] got [", dim(-2, 5), "]"
    stop 1
end if
end program t
