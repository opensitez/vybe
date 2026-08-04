! vybe-test: fortran/fortran2018/assumed_rank_with_select_rank
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

module ar_mod
    implicit none
contains
    subroutine process(x)
        real, intent(in) :: x(..)
        select rank(x)
        rank(0)
            print *, 'scalar', x
        rank(1)
            print *, 'vector of size', size(x)
        rank(2)
            print *, 'matrix', size(x,1), 'x', size(x,2)
        rank default
            print *, 'rank', rank(x)
        end select
    end subroutine process
end module ar_mod

program test
    use ar_mod
    real :: v(5) = [1., 2., 3., 4., 5.]
    call process(v)
end program test
