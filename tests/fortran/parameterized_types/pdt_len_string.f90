! vybe-test: fortran/parameterized_types/pdt_len_string
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: BoundedStr(maxlen)
        integer, len :: maxlen
        character(len=maxlen) :: value
    end type BoundedStr
    type(BoundedStr(20)) :: s
    s%value = 'hello'
    print *, trim(s%value)
end program test
