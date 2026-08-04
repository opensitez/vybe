! vybe-test: fortran/reshape_pad_extended/reshape_shape_product_12
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(12) = [(i, i = 1, 12)]
integer :: m(3,4)
m = reshape(a, [3, 4])
if ((m(3,4)) /= 12) then
    print *, "FAIL: want [12] got [", m(3,4), "]"
    stop 1
end if
if ((sum(m)) /= 78) then
    print *, "FAIL: want [78] got [", sum(m), "]"
    stop 1
end if
end program t
