! vybe-test: fortran/reshape_pad_extended/reshape_3d_order_c_differs
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: mf(2,2,1), mc(2,2,1)
mf = reshape(a, [2, 2, 1])
mc = reshape(a, [2, 2, 1], order=[3, 2, 1])
if ((mf(2,1,1)) /= 2) then
    print *, "FAIL: want [2] got [", mf(2,1,1), "]"
    stop 1
end if
if ((mc(2,1,1)) /= 3) then
    print *, "FAIL: want [3] got [", mc(2,1,1), "]"
    stop 1
end if
end program t
