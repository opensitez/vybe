! vybe-test: fortran/modulo_dim_sign_extended/merge_real_abs_via_sign
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: x
x = merge(-6.0, 6.0, .false.)
if ((nint(x)) /= 6) then
    print *, "FAIL: want [6] got [", nint(x), "]"
    stop 1
end if
end program t
