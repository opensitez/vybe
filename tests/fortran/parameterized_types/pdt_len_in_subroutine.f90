! vybe-test: fortran/parameterized_types/pdt_len_in_subroutine
! origin: languages/fortran/tests/fortran/test_parameterized_types.rs

program test
    type :: Vec(n)
        integer, len :: n
        real :: data(n)
    end type Vec
    type(Vec(4)) :: v
    v%data = [1.0, 2.0, 3.0, 4.0]
    call show_first(v)
contains
    subroutine show_first(v)
        type(Vec(*)), intent(in) :: v
        print *, v%data(1)
    end subroutine show_first
end program test
