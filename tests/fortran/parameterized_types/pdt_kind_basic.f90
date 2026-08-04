! vybe-test: fortran/parameterized_types/pdt_kind_basic
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: TypedNum(k)
        integer, kind :: k
        real(k) :: value
    end type TypedNum
    type(TypedNum(4)) :: f
    type(TypedNum(8)) :: d
    f%value = 3.14_4
    d%value = 3.14159265358979_8
    print *, f%value
    print *, d%value
end program test
