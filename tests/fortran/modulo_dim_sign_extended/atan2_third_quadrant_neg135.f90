! vybe-test: fortran/modulo_dim_sign_extended/atan2_third_quadrant_neg135
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(-1.0, -1.0)*180/3.14159265)) /= -135) then
    print *, "FAIL: want [-135] got [", nint(atan2(-1.0, -1.0)*180/3.14159265), "]"
    stop 1
end if
end program t
