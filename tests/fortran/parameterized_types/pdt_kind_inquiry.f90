! vybe-test: fortran/parameterized_types/pdt_kind_inquiry
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Num(k)
        integer, kind :: k
        real(k) :: v
    end type Num
    type(Num(8)) :: x
    x%v = 1.0_8
    print *, x%k
end program test
