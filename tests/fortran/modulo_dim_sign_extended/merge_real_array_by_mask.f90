! vybe-test: fortran/modulo_dim_sign_extended/merge_real_array_by_mask
! origin: languages/fortran/tests/fortran/test_modulo_dim_sign_extended.rs
program t
real :: a(3)=[1.5,2.5,3.5]
real :: b(3)=[9.0,8.0,7.0]
real :: c(3)
logical :: m(3)=[.true.,.false.,.true.]
c = merge(a, b, m)
if ((nint(c(1)*10)) /= 15) then
    print *, "FAIL: want [15] got [", nint(c(1)*10), "]"
    stop 1
end if
if ((nint(c(2)*10)) /= 80) then
    print *, "FAIL: want [80] got [", nint(c(2)*10), "]"
    stop 1
end if
if ((nint(c(3)*10)) /= 35) then
    print *, "FAIL: want [35] got [", nint(c(3)*10), "]"
    stop 1
end if
end program t
