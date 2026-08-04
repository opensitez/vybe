! vybe-test: fortran/reshape_pad_extended/reshape_3d_order_c_first_slice
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
integer :: m(2,2,2)
m = reshape(a, [2, 2, 2], order='C')
if ((m(1,1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1,1), "]"
    stop 1
end if
if ((m(1,1,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(1,1,2), "]"
    stop 1
end if
if ((m(2,2,2)) /= 8) then
    print *, "FAIL: want [8] got [", m(2,2,2), "]"
    stop 1
end if
end program t
