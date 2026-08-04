! vybe-test: fortran/pointer_alloc_extended/alloc_3d_shape_product
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: cube(:,:,:)
allocate(cube(2, 3, 4))
if ((size(cube)) /= 24) then
    print *, "FAIL: want [24] got [", size(cube), "]"
    stop 1
end if
if ((size(cube, 1)) /= 2) then
    print *, "FAIL: want [2] got [", size(cube, 1), "]"
    stop 1
end if
if ((size(cube, 2)) /= 3) then
    print *, "FAIL: want [3] got [", size(cube, 2), "]"
    stop 1
end if
if ((size(cube, 3)) /= 4) then
    print *, "FAIL: want [4] got [", size(cube, 3), "]"
    stop 1
end if
deallocate(cube)
end program t
