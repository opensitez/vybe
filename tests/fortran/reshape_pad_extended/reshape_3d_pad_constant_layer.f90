! vybe-test: fortran/reshape_pad_extended/reshape_3d_pad_constant_layer
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [1, 2, 3]
integer :: m(1,1,4)
m = reshape(a, [1, 1, 4], pad=9)
if ((m(1,1,4)) /= 9) then
    print *, "FAIL: want [9] got [", m(1,1,4), "]"
    stop 1
end if
end program t
