! vybe-test: fortran/select_type_polymorphic_matching/same_type_as
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: A
        integer :: x = 1
    end type A
    type(A) :: obj1, obj2
    print *, same_type_as(obj1, obj2)
end program test
