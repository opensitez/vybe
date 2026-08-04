! vybe-test: fortran/parameterized_types/pdt_len_and_kind
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Precision(k, n)
        integer, kind :: k
        integer, len :: n
        real(k) :: data(n)
    end type Precision
    type(Precision(8, 5)) :: p
    p%data = 1.0_8
    p%data(3) = 3.14159265_8
    print *, p%data(3)
end program test
