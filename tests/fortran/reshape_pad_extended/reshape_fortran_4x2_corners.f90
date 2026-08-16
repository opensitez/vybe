! vybe-test: fortran/reshape_pad_extended/reshape_fortran_4x2_corners
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]
integer :: m(4,2)
m = reshape(a, [4, 2])
if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
if ((m(4,1)) /= 4) then
    print *, "FAIL: want [4] got [", m(4,1), "]"
    stop 1
end if
if ((m(4,2)) /= 8) then
    print *, "FAIL: want [8] got [", m(4,2), "]"
    stop 1
end if
end program t
