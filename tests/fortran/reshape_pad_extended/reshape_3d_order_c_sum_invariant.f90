! vybe-test: fortran/reshape_pad_extended/reshape_3d_order_c_sum_invariant
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(12) = [(i, i = 1, 12)]
integer :: m(2,3,2)
m = reshape(a, [2, 3, 2], order='C')
if ((sum(m)) /= 78) then
    print *, "FAIL: want [78] got [", sum(m), "]"
    stop 1
end if
end program t
