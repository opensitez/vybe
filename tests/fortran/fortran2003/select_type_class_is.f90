! vybe-test: fortran/fortran2003/select_type_class_is
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: A
        integer :: x = 1
    end type A
    type, extends(A) :: B
        integer :: y = 2
    end type B
    class(A), allocatable :: obj
    allocate(B :: obj)
    select type(obj)
    class is (B)
        print *, obj%y
    type is (A)
        print *, obj%x
    end select
end program test
