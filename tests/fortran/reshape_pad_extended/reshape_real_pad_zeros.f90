! vybe-test: fortran/reshape_pad_extended/reshape_real_pad_zeros
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
real :: a(2) = [1.0, 2.0]
real :: m(3)
m = reshape(a, [3], pad=[0.0])
if ((int(m(3) * 10)) /= 0) then
    print *, "FAIL: want [0] got [", int(m(3) * 10), "]"
    stop 1
end if
end program t
