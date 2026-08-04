! vybe-test: fortran/reshape_pad_extended/reshape_fortran_2x4_second_row
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(8) = [1, 2, 3, 4, 5, 6, 7, 8]
integer :: m(2,4)
m = reshape(a, [2, 4])
if ((m(1,2)) /= 3) then
    print *, "FAIL: want [3] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", m(2,2), "]"
    stop 1
end if
if ((m(2,4)) /= 8) then
    print *, "FAIL: want [8] got [", m(2,4), "]"
    stop 1
end if
end program t
