! vybe-test: fortran/allocatable_components/compile_alloc_comp_deferred_shape_rank
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Flex
        integer, allocatable :: buf(:)
    end type Flex
    type(Flex) :: f
    allocate(f%buf(0:2))
    f%buf = [1, 2, 3]
    print *, f%buf(0) + f%buf(2)
end program t
