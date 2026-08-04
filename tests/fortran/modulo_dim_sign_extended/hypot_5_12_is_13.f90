! vybe-test: fortran/modulo_dim_sign_extended/hypot_5_12_is_13
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(5.0, 12.0))) /= 13) then
    print *, "FAIL: want [13] got [", nint(hypot(5.0, 12.0)), "]"
    stop 1
end if
end program t
