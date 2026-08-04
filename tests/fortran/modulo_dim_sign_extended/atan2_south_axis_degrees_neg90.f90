! vybe-test: fortran/modulo_dim_sign_extended/atan2_south_axis_degrees_neg90
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(-1.0, 0.0)*180/3.14159265)) /= -90) then
    print *, "FAIL: want [-90] got [", nint(atan2(-1.0, 0.0)*180/3.14159265), "]"
    stop 1
end if
end program t
