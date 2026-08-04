! vybe-test: fortran/reshape_pad_extended/reshape_real_2x2_fractions
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
real :: a(4) = [0.5, 1.5, 2.5, 3.5]
real :: m(2,2)
m = reshape(a, [2, 2])
if ((int(sum(m) * 10)) /= 80) then
    print *, "FAIL: want [80] got [", int(sum(m) * 10), "]"
    stop 1
end if
end program t
