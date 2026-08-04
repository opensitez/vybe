! vybe-test: fortran/parameterized_types/pdt_deferred_len
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: DynVec(n)
        integer, len :: n
        real :: data(n)
    end type DynVec
    type(DynVec(:)), allocatable :: v
    allocate(DynVec(10) :: v)
    v%data = 0.0
    v%data(5) = 42.0
    print *, v%data(5)
    deallocate(v)
end program test
