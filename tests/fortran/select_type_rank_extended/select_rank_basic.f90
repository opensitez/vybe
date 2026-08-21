! vybe-test: fortran/select_type_rank_extended/select_rank_basic
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program test
    call handle([1, 2, 3])
contains
    subroutine handle(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(1)
            print *, 'rank-1 array, size =', size(x)
        rank default
            print *, 'other rank:', rank(x)
        end select
    end subroutine handle
end program test
