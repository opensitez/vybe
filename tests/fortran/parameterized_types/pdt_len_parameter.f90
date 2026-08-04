! vybe-test: fortran/parameterized_types/pdt_len_parameter
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Matrix(m, n)
        integer, len :: m, n
        real :: data(m, n)
    end type Matrix
    type(Matrix(3,3)) :: mat
    mat%data = 0.0
    mat%data(2,2) = 99.0
    print *, mat%data(2,2)
end program test
