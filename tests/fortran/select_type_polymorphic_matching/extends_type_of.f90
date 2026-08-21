! vybe-test: fortran/select_type_polymorphic_matching/extends_type_of
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Base
        integer :: x = 0
    end type Base
    type, extends(Base) :: Child
        integer :: y = 1
    end type Child
    type(Base) :: b
    type(Child) :: c
    print *, extends_type_of(c, b)
end program test
