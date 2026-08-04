! vybe-test: fortran/parameterized_types/pdt_len_inquiry
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Str(n)
        integer, len :: n
        character(n) :: s
    end type Str
    type(Str(15)) :: t
    t%s = 'hello'
    print *, t%n
end program test
