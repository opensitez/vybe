! vybe-test: fortran/where_advanced/where_in_subroutine
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(5) = [3, -1, 5, -2, 4]
    call clamp_negatives(a)
    print *, a(2)
    print *, a(4)
contains
    subroutine clamp_negatives(x)
        integer, intent(inout) :: x(:)
        where (x < 0)
            x = 0
        end where
    end subroutine clamp_negatives
end program test
