! vybe-test: fortran/fortran2003/type_bound_procedure_allows_keyword_binding_name
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: stats_result
        integer :: n = 12
    contains
        procedure :: print => print_stats
    end type stats_result

    type(stats_result) :: stats
    call stats%print()
contains
    subroutine print_stats(self)
        class(stats_result), intent(in) :: self
        print *, "n =", self%n
    end subroutine print_stats
end program test
