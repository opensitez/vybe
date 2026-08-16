! vybe-test: fortran/reshape_pad_extended/reshape_flatten_order_c
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2,2) = reshape([1, 2, 3, 4], [2, 2])
integer :: b(4)
b = reshape(a, [4], order=[1])
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 4) then
    print *, "FAIL: want [4] got [", b(4), "]"
    stop 1
end if
end program t
