! vybe-test: fortran/reshape_pad_extended/reshape_pad_integer_nine_fill
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [1, 2, 3]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[9])
if ((m(1,2)) /= 3) then
    print *, "FAIL: want [3] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,1)) /= 2) then
    print *, "FAIL: want [2] got [", m(2,1), "]"
    stop 1
end if
if ((m(2,2)) /= 9) then
    print *, "FAIL: want [9] got [", m(2,2), "]"
    stop 1
end if
end program t
