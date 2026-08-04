! vybe-test: fortran/modulo_dim_sign_extended/hypot_both_zero
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(0.0, 0.0))) /= 0) then
    print *, "FAIL: want [0] got [", nint(hypot(0.0, 0.0)), "]"
    stop 1
end if
end program t
