! vybe-test: fortran/pointers/allocatable_2d_in_type_runtime
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    type :: NeighborGrid
        integer, allocatable :: cells(:,:)
    end type NeighborGrid
    type(NeighborGrid) :: grid

    allocate(grid%cells(2,3))
    grid%cells(1,1) = 5
    grid%cells(2,3) = 8
    if ((grid%cells(1,1)) /= 5) then
    print *, "FAIL: want [5] got [", grid%cells(1,1), "]"
    stop 1
end if
    if ((grid%cells(2,3)) /= 8) then
    print *, "FAIL: want [8] got [", grid%cells(2,3), "]"
    stop 1
end if
    deallocate(grid%cells)
end program test
