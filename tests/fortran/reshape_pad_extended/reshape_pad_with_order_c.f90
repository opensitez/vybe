! vybe-test: fortran/reshape_pad_extended/reshape_pad_with_order_c
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [1, 2, 3]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[0], order=[2, 1])
if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
if ((m(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,2), "]"
    stop 1
end if
end program t
