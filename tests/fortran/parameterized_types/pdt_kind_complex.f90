! vybe-test: fortran/parameterized_types/pdt_kind_complex
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: TypedComplex(k)
        integer, kind :: k
        complex(k) :: value
    end type TypedComplex
    type(TypedComplex(4)) :: c
    c%value = (1.0_4, 2.0_4)
    print *, real(c%value)
end program test
