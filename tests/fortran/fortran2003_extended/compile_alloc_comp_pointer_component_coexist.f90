! vybe-test: fortran/fortran2003_extended/compile_alloc_comp_pointer_component_coexist
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Mix
        integer, allocatable :: owned(:)
        integer, pointer :: view(:) => null()
    end type Mix
    type(Mix) :: m
    allocate(m%owned(2))
    m%owned = [4, 5]
    m%view => m%owned
    print *, m%view(1)
end program t
