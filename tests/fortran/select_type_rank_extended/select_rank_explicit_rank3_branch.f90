! vybe-test: fortran/select_type_rank_extended/select_rank_explicit_rank3_branch
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs

program t
    call inspect(reshape([(i, i = 1, 24)], [2, 3, 4]))
contains
    subroutine inspect(x)
        integer, intent(in) :: x(..)
        select rank(x)
        rank(3)
            print *, size(x, 1), size(x, 2), size(x, 3)
        rank default
            print *, rank(x)
        end select
    end subroutine inspect
end program t
