! vybe-test: fortran/reshape_pad_extended/reshape_pad_integer_zero_fill
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: m(2,3)
m = reshape(a, [2, 3], pad=[0])
if ((m(1,3)) /= 0) then
    print *, "FAIL: want [0] got [", m(1,3), "]"
    stop 1
end if
if ((m(2,3)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,3), "]"
    stop 1
end if
if ((sum(m)) /= 10) then
    print *, "FAIL: want [10] got [", sum(m), "]"
    stop 1
end if
end program t
