! vybe-test: fortran/modulo_dim_sign_extended/merge_real_scalar_false_branch
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: x
x = merge(3.5, 7.5, .false.)
if ((nint(x*10)) /= 75) then
    print *, "FAIL: want [75] got [", nint(x*10), "]"
    stop 1
end if
end program t
