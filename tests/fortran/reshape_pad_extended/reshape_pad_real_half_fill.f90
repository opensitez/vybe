! vybe-test: fortran/reshape_pad_extended/reshape_pad_real_half_fill
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
real :: a(2) = [1.5, 2.5]
real :: m(2,2)
m = reshape(a, [2, 2], pad=[0.0])
! Fortran fills COLUMN-major, so the source occupies m(1,1) and m(2,1) and the
! PADDED entries are the second column — m(1,2) and m(2,2).
if ((int(m(1,2) * 10)) /= 0) then
    print *, "FAIL: want [0] got [", int(m(1,2) * 10), "]"
    stop 1
end if
if ((int(m(2,2) * 10)) /= 0) then
    print *, "FAIL: want [0] got [", int(m(2,2) * 10), "]"
    stop 1
end if
end program t
