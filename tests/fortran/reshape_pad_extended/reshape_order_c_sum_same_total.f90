! vybe-test: fortran/reshape_pad_extended/reshape_order_c_sum_same_total
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
integer :: mf(2,3), mc(2,3)
mf = reshape(a, [2, 3])
mc = reshape(a, [2, 3], order=[2, 1])
if ((sum(mf)) /= 21) then
    print *, "FAIL: want [21] got [", sum(mf), "]"
    stop 1
end if
if ((sum(mc)) /= 21) then
    print *, "FAIL: want [21] got [", sum(mc), "]"
    stop 1
end if
end program t
