! vybe-test: fortran/reshape_pad_extended/reshape_pad_5x1_from_two
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2) = [8, 9]
integer :: m(5,1)
m = reshape(a, [5, 1], pad=0)
if ((m(1,1)) /= 8) then
    print *, "FAIL: want [8] got [", m(1,1), "]"
    stop 1
end if
if ((m(5,1)) /= 0) then
    print *, "FAIL: want [0] got [", m(5,1), "]"
    stop 1
end if
end program t
