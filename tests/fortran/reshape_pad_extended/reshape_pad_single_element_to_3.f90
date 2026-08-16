! vybe-test: fortran/reshape_pad_extended/reshape_pad_single_element_to_3
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(1) = [42]
integer :: m(3)
m = reshape(a, [3], pad=[-1])
if ((m(1)) /= 42) then
    print *, "FAIL: want [42] got [", m(1), "]"
    stop 1
end if
if ((m(2)) /= -1) then
    print *, "FAIL: want [-1] got [", m(2), "]"
    stop 1
end if
if ((m(3)) /= -1) then
    print *, "FAIL: want [-1] got [", m(3), "]"
    stop 1
end if
end program t
