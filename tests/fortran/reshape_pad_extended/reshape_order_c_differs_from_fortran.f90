! vybe-test: fortran/reshape_pad_extended/reshape_order_c_differs_from_fortran
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: mf(2,2), mc(2,2)
mf = reshape(a, [2, 2])
mc = reshape(a, [2, 2], order='C')
if ((mf(1,2)) /= 3) then
    print *, "FAIL: want [3] got [", mf(1,2), "]"
    stop 1
end if
if ((mc(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", mc(1,2), "]"
    stop 1
end if
end program t
