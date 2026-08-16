! vybe-test: fortran/reshape_pad_extended/reshape_shape_product_24_with_pad
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
integer :: m(2,3,4)
m = reshape(a, [2, 3, 4], pad=[0])
if ((sum(m)) /= 55) then
    print *, "FAIL: want [55] got [", sum(m), "]"
    stop 1
end if
end program t
