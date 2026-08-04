! vybe-test: fortran/reshape_pad_extended/reshape_1d_to_1d_identity
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(5) = [5, 4, 3, 2, 1]
integer :: b(5)
b = reshape(a, [5])
if ((b(1)) /= 5) then
    print *, "FAIL: want [5] got [", b(1), "]"
    stop 1
end if
if ((b(5)) /= 1) then
    print *, "FAIL: want [1] got [", b(5), "]"
    stop 1
end if
end program t
