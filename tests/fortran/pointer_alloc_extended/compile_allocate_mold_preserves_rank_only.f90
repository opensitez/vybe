! vybe-test: fortran/pointer_alloc_extended/compile_allocate_mold_preserves_rank_only
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

program t
    integer, allocatable :: pattern(:,:), blank(:,:)
    allocate(pattern(2, 3))
    pattern = 0
    allocate(blank, mold=pattern)
    print *, size(blank, 1)
    print *, size(blank, 2)
    deallocate(pattern, blank)
end program t
