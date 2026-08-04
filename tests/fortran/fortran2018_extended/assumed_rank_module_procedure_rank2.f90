! vybe-test: fortran/fortran2018_extended/assumed_rank_module_procedure_rank2
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

module ranks
    implicit none
contains
    subroutine rows(x)
        real, intent(in) :: x(..)
        select rank(x)
        rank(2)
            print *, size(x, 1)
        rank default
            print *, 0
        end select
    end subroutine rows
end module ranks

program t
    use ranks
    real :: grid(4, 3)
    call rows(grid)
end program t
