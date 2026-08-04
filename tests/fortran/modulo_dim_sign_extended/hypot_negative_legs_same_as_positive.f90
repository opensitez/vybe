! vybe-test: fortran/modulo_dim_sign_extended/hypot_negative_legs_same_as_positive
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(hypot(-3.0, -4.0))) /= 5) then
    print *, "FAIL: want [5] got [", nint(hypot(-3.0, -4.0)), "]"
    stop 1
end if
end program t
