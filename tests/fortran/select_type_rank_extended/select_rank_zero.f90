! vybe-test: fortran/select_type_rank_extended/select_rank_zero
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    integer :: s = 99
    call inspect(s)
contains
    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(0)
            print *, 'scalar =', x
        rank(1)
            print *, 'vector'
        end select
    end subroutine inspect
end program test
