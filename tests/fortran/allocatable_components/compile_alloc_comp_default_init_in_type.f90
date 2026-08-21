! vybe-test: fortran/allocatable_components/compile_alloc_comp_default_init_in_type
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Defaults
        integer, allocatable :: vals(:)
    end type Defaults
    type(Defaults) :: d
    d%vals = [9]
    print *, d%vals(1)
end program t
