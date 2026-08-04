! vybe-test: fortran/reshape_pad_extended/reshape_order_c_3x3_center
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(9) = [(i, i = 1, 9)]
integer :: m(3,3)
m = reshape(a, [3, 3], order='C')
if ((m(2,2)) /= 5) then
    print *, "FAIL: want [5] got [", m(2,2), "]"
    stop 1
end if
if ((m(3,3)) /= 9) then
    print *, "FAIL: want [9] got [", m(3,3), "]"
    stop 1
end if
end program t
