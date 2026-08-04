! vybe-test: fortran/reshape_pad_extended/reshape_flatten_2d_to_1d
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2,3) = reshape([1, 2, 3, 4, 5, 6], [2, 3])
integer :: b(6)
b = reshape(a, [6])
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(6)) /= 6) then
    print *, "FAIL: want [6] got [", b(6), "]"
    stop 1
end if
end program t
