! vybe-test: fortran/reshape_pad_extended/reshape_pad_2x2_from_three_elements
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [10, 20, 30]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[0])
if ((m(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,2)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,2), "]"
    stop 1
end if
end program t
