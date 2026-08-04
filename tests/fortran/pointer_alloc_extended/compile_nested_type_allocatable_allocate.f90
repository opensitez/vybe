! vybe-test: fortran/pointer_alloc_extended/compile_nested_type_allocatable_allocate
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs

type :: Inner
    real, allocatable :: coeffs(:)
end type Inner

type :: Outer
    type(Inner) :: layer
end type Outer

program t
    type(Outer) :: obj
    allocate(obj%layer%coeffs(3))
    obj%layer%coeffs = [1.0, 2.0, 3.0]
    print *, obj%layer%coeffs(2)
    deallocate(obj%layer%coeffs)
end program t
