! vybe-test: fortran/modulo_dim_sign_extended/merge_real_sum_selected
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: a(3)=[0.5,1.5,2.5]
real :: b(3)=[5.0,4.0,3.0]
logical :: m(3)=[.true.,.false.,.true.]
if ((nint(sum(merge(a,b,m))*10)) /= 90) then
    print *, "FAIL: want [90] got [", nint(sum(merge(a,b,m))*10), "]"
    stop 1
end if
end program t
