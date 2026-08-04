! vybe-test: fortran/reshape_pad_extended/reshape_3d_3x1x2
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [10, 20, 30, 40, 50, 60]
integer :: m(3,1,2)
m = reshape(a, [3, 1, 2])
if ((m(1,1,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,1,1), "]"
    stop 1
end if
if ((m(3,1,2)) /= 60) then
    print *, "FAIL: want [60] got [", m(3,1,2), "]"
    stop 1
end if
end program t
