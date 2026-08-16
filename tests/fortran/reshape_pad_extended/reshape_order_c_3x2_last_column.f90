! vybe-test: fortran/reshape_pad_extended/reshape_order_c_3x2_last_column
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
integer :: m(3,2)
m = reshape(a, [3, 2], order=[2, 1])
if ((m(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(1,2), "]"
    stop 1
end if
if ((m(3,2)) /= 6) then
    print *, "FAIL: want [6] got [", m(3,2), "]"
    stop 1
end if
end program t
