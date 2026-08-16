! vybe-test: fortran/reshape_pad_extended/reshape_pad_negative_pad_value
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2) = [5, 10]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[-99])
if ((m(1,2)) /= -99) then
    print *, "FAIL: want [-99] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(2,1), "]"
    stop 1
end if
end program t
