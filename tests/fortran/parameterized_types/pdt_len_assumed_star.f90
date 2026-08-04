! vybe-test: fortran/parameterized_types/pdt_len_assumed_star
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Buf(n)
        integer, len :: n
        integer :: items(n)
    end type Buf
    type(Buf(10)) :: b
    b%items = [(i, i=1,10)]
    call process(b)
contains
    subroutine process(x)
        type(Buf(*)), intent(in) :: x
        print *, x%items(1)
        print *, size(x%items)
    end subroutine process
end program test
