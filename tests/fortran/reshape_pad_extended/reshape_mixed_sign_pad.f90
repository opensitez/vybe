! vybe-test: fortran/reshape_pad_extended/reshape_mixed_sign_pad
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(3) = [1, -1, 1]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[-1])
if ((m(2,2)) /= -1) then
    print *, "FAIL: want [-1] got [", m(2,2), "]"
    stop 1
end if
end program t
