! vybe-test: fortran/reshape_pad_extended/reshape_3d_2x3x1_column
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
integer :: m(2,3,1)
m = reshape(a, [2, 3, 1])
if ((m(2,3,1)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,3,1), "]"
    stop 1
end if
if ((sum(m)) /= 21) then
    print *, "FAIL: want [21] got [", sum(m), "]"
    stop 1
end if
end program t
