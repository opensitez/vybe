! vybe-test: fortran/modulo_dim_sign_extended/merge_real_with_logical_from_compare
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: a=2.0, b=5.0
if ((nint(merge(a, b, a<b)*10)) /= 20) then
    print *, "FAIL: want [20] got [", nint(merge(a, b, a<b)*10), "]"
    stop 1
end if
end program t
