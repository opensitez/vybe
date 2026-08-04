! vybe-test: fortran/modulo_dim_sign_extended/atan2_fourth_quadrant_neg45
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
if ((nint(atan2(-1.0, 1.0)*180/3.14159265)) /= -45) then
    print *, "FAIL: want [-45] got [", nint(atan2(-1.0, 1.0)*180/3.14159265), "]"
    stop 1
end if
end program t
