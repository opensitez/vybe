! vybe-test: fortran/modulo_dim_sign_extended/merge_real_2d_slice
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: a(2,2)=reshape([1.0,2.0,3.0,4.0],[2,2])
real :: b(2,2)=reshape([10.0,20.0,30.0,40.0],[2,2])
real :: c(2,2)
c = merge(a, b, a<3.0)
if ((nint(c(1,1)*10)) /= 10) then
    print *, "FAIL: want [10] got [", nint(c(1,1)*10), "]"
    stop 1
end if
if ((nint(c(2,2)*10)) /= 400) then
    print *, "FAIL: want [400] got [", nint(c(2,2)*10), "]"
    stop 1
end if
end program t
