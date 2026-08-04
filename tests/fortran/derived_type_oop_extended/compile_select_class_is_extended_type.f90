! vybe-test: fortran/derived_type_oop_extended/compile_select_class_is_extended_type
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: A
        integer :: x = 1
    end type A
    type, extends(A) :: B
        integer :: y = 7
    end type B
    class(A), allocatable :: obj
    allocate(B :: obj)
    select type(obj)
    class is (B)
        print *, obj%y
    type is (A)
        print *, obj%x
    end select
end program t
