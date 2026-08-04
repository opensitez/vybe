! vybe-test: fortran/parameterized_types/pdt_default_len
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: DefaultVec(n)
        integer, len :: n = 10
        real :: data(n)
    end type DefaultVec
    type(DefaultVec()) :: v
    v%data = 0.0
    print *, size(v%data)
end program test
