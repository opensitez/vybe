! vybe-test: fortran/parameterized_types/pdt_len_basic
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: FixedVec(n)
        integer, len :: n
        real :: data(n)
    end type FixedVec
    type(FixedVec(5)) :: v
    v%data = 0.0
    v%data(1) = 1.0
    print *, v%data(1)
end program test
