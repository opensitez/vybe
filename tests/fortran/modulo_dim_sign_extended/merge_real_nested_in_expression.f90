! vybe-test: fortran/modulo_dim_sign_extended/merge_real_nested_in_expression
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: x
x = merge(merge(1.0,2.0,.true.), merge(3.0,4.0,.false.), .true.)
if ((nint(x*10)) /= 10) then
    print *, "FAIL: want [10] got [", nint(x*10), "]"
    stop 1
end if
end program t
