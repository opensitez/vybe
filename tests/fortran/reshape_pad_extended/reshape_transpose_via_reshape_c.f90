! vybe-test: fortran/reshape_pad_extended/reshape_transpose_via_reshape_c
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2,3) = reshape([1, 4, 2, 5, 3, 6], [2, 3])
integer :: flat(6), back(3,2)
flat = reshape(a, [6], order='C')
back = reshape(flat, [3, 2], order='C')
if ((back(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", back(1,1), "]"
    stop 1
end if
if ((back(3,2)) /= 6) then
    print *, "FAIL: want [6] got [", back(3,2), "]"
    stop 1
end if
end program t
