! vybe-test: fortran/modulo_dim_sign_extended/hypot_scaled_triangle
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(30.0, 40.0))) /= 50) then
    print *, "FAIL: want [50] got [", nint(hypot(30.0, 40.0)), "]"
    stop 1
end if
end program t
