! vybe-test: fortran/reshape_pad_extended/reshape_order_c_4x1_column
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [9, 8, 7, 6]
integer :: m(4,1)
m = reshape(a, [4, 1], order='C')
if ((m(1,1)) /= 9) then
    print *, "FAIL: want [9] got [", m(1,1), "]"
    stop 1
end if
if ((m(4,1)) /= 6) then
    print *, "FAIL: want [6] got [", m(4,1), "]"
    stop 1
end if
end program t
