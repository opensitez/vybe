! vybe-test: fortran/parameterized_types/pdt_kind_integer
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: TypedInt(k)
        integer, kind :: k
        integer(k) :: value
    end type TypedInt
    type(TypedInt(8)) :: big
    big%value = 1000000000_8
    print *, big%value
end program test
