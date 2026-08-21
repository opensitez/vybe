! vybe-test: fortran/elemental_procedure_special_cases/impure_elemental_basic
! origin: languages/fortran/tests/fortran/test_fortran2008.rs

program test
    integer :: a(3) = [1, 2, 3]
    call print_elem(a)
contains
    impure elemental subroutine print_elem(x)
        integer, intent(in) :: x
        print *, x
    end subroutine print_elem
end program test
