! vybe-test: fortran/parameterized_types/pdt_allocatable_component
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: DynMat(m, n)
        integer, len :: m, n
        real, allocatable :: data(:,:)
    end type DynMat
    type(DynMat(3,4)) :: mat
    allocate(mat%data(mat%m, mat%n))
    mat%data = 0.0
    mat%data(2,3) = 7.0
    print *, mat%data(2,3)
    deallocate(mat%data)
end program test
