! vybe-test: fortran/modulo_dim_sign_extended/atan2_west_axis_degrees_180
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(0.0, -1.0)*180/3.14159265)) /= 180) then
    print *, "FAIL: want [180] got [", nint(atan2(0.0, -1.0)*180/3.14159265), "]"
    stop 1
end if
end program t
