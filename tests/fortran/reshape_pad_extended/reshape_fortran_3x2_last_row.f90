! vybe-test: fortran/reshape_pad_extended/reshape_fortran_3x2_last_row
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
integer :: m(3,2)
m = reshape(a, [3, 2])
if ((m(3,1)) /= 5) then
    print *, "FAIL: want [5] got [", m(3,1), "]"
    stop 1
end if
if ((m(3,2)) /= 6) then
    print *, "FAIL: want [6] got [", m(3,2), "]"
    stop 1
end if
if ((sum(m)) /= 21) then
    print *, "FAIL: want [21] got [", sum(m), "]"
    stop 1
end if
end program t
