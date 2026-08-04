! vybe-test: fortran/where_advanced/storage_size_derived_type
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    type :: Pair
        integer :: x, y
    end type Pair
    type(Pair) :: p
    print *, storage_size(p)
end program test
