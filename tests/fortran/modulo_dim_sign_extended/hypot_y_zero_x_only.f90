! vybe-test: fortran/modulo_dim_sign_extended/hypot_y_zero_x_only
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(7.0, 0.0))) /= 7) then
    print *, "FAIL: want [7] got [", nint(hypot(7.0, 0.0)), "]"
    stop 1
end if
end program t
