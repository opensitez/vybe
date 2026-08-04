! vybe-test: fortran/modulo_dim_sign_extended/atan2_east_axis_degrees_zero
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(0.0, 1.0)*180/3.14159265)) /= 0) then
    print *, "FAIL: want [0] got [", nint(atan2(0.0, 1.0)*180/3.14159265), "]"
    stop 1
end if
end program t
