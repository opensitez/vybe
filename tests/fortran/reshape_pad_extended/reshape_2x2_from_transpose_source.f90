! vybe-test: fortran/reshape_pad_extended/reshape_2x2_from_transpose_source
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2,2) = reshape([1, 3, 2, 4], [2, 2])
integer :: b(2,2)
b = reshape(a, [2, 2])
if ((b(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1,1), "]"
    stop 1
end if
if ((b(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", b(2,2), "]"
    stop 1
end if
end program t
