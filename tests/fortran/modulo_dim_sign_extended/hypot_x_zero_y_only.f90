! vybe-test: fortran/modulo_dim_sign_extended/hypot_x_zero_y_only
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(0.0, 9.0))) /= 9) then
    print *, "FAIL: want [9] got [", nint(hypot(0.0, 9.0)), "]"
    stop 1
end if
end program t
