! vybe-test: fortran/modulo_dim_sign_extended/hypot_8_15_is_17
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(8.0, 15.0))) /= 17) then
    print *, "FAIL: want [17] got [", nint(hypot(8.0, 15.0)), "]"
    stop 1
end if
end program t
