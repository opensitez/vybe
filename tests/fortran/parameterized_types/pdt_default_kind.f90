! vybe-test: fortran/parameterized_types/pdt_default_kind
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: DefaultReal(k)
        integer, kind :: k = 4
        real(k) :: x
    end type DefaultReal
    type(DefaultReal()) :: r
    r%x = 1.0
    print *, r%x
end program test
