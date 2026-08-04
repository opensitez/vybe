! vybe-test: fortran/reshape_pad_extended/reshape_3d_fortran_2x2x2_corner
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(8) = [(i, i = 1, 8)]
integer :: m(2,2,2)
m = reshape(a, [2, 2, 2])
if ((m(1,1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1,1), "]"
    stop 1
end if
if ((m(2,2,2)) /= 8) then
    print *, "FAIL: want [8] got [", m(2,2,2), "]"
    stop 1
end if
if ((sum(m)) /= 36) then
    print *, "FAIL: want [36] got [", sum(m), "]"
    stop 1
end if
end program t
