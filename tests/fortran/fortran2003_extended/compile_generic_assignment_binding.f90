! vybe-test: fortran/fortran2003_extended/compile_generic_assignment_binding
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    type :: Bag
        integer, allocatable :: items(:)
    contains
        procedure :: assign_from
        generic :: assignment(=) => assign_from
    end type Bag
    type(Bag) :: a, b
    a%items = [1, 2]
    b = a
    print *, size(b%items)
contains
    subroutine assign_from(lhs, rhs)
        class(Bag), intent(out) :: lhs
        type(Bag), intent(in) :: rhs
        lhs%items = rhs%items
    end subroutine assign_from
end program t
