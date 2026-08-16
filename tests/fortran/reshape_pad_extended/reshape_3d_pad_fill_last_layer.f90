! vybe-test: fortran/reshape_pad_extended/reshape_3d_pad_fill_last_layer
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: m(2,2,2)
m = reshape(a, [2, 2, 2], pad=[0])
if ((m(2,2,2)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,2,2), "]"
    stop 1
end if
if ((count(m == 0)) /= 3) then
    print *, "FAIL: want [3] got [", count(m == 0), "]"
    stop 1
end if
end program t
