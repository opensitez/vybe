! vybe-test: fortran/modulo_dim_sign_extended/atan2_neg12_5_obtuse
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(5.0, -12.0)*180/3.14159265)) /= 157) then
    print *, "FAIL: want [157] got [", nint(atan2(5.0, -12.0)*180/3.14159265), "]"
    stop 1
end if
end program t
