! vybe-test: fortran/reshape_pad_extended/reshape_pad_and_order_c_combined
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [10, 20, 30]
integer :: m(2,3)
m = reshape(a, [2, 3], pad=1, order='C')
if ((m(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,3)) /= 1) then
    print *, "FAIL: want [1] got [", m(2,3), "]"
    stop 1
end if
end program t
