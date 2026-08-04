! vybe-test: fortran/modulo_dim_sign_extended/merge_real_with_negative_values
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: x
x = merge(-2.5, 4.0, .false.)
if ((nint(x*10)) /= 40) then
    print *, "FAIL: want [40] got [", nint(x*10), "]"
    stop 1
end if
end program t
