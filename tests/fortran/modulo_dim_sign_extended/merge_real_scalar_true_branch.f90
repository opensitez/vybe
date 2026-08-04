! vybe-test: fortran/modulo_dim_sign_extended/merge_real_scalar_true_branch
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: x
x = merge(3.5, 7.5, .true.)
if ((nint(x*10)) /= 35) then
    print *, "FAIL: want [35] got [", nint(x*10), "]"
    stop 1
end if
end program t
