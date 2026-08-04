! vybe-test: fortran/modulo_dim_sign_extended/atan2_3_4_triangle_degrees
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(3.0, 4.0)*180/3.14159265)) /= 37) then
    print *, "FAIL: want [37] got [", nint(atan2(3.0, 4.0)*180/3.14159265), "]"
    stop 1
end if
end program t
