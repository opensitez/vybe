! vybe-test: fortran/pointer_alloc_extended/compile_allocate_source_from_expression
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

program t
    integer, allocatable :: base(:), copy(:)
    base = [2, 4, 6, 8]
    allocate(copy, source=base + 1)
    print *, copy(1)
    print *, copy(4)
    deallocate(base, copy)
end program t
