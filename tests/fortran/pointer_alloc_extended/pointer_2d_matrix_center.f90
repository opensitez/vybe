! vybe-test: fortran/pointer_alloc_extended/pointer_2d_matrix_center
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: mat(2, 2)
integer, pointer :: view(:,:)
mat = reshape([1, 2, 3, 4], [2, 2])
view => mat
if ((view(2, 1)) /= 3) then
    print *, "FAIL: want [3] got [", view(2, 1), "]"
    stop 1
end if
end program t
