! vybe-test: fortran/reshape_pad_extended/reshape_pad_3x3_from_five
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: m(3,3)
m = reshape(a, [3, 3], pad=[7])
if ((m(3,3)) /= 7) then
    print *, "FAIL: want [7] got [", m(3,3), "]"
    stop 1
end if
if ((count(m == 7)) /= 4) then
    print *, "FAIL: want [4] got [", count(m == 7), "]"
    stop 1
end if
end program t
