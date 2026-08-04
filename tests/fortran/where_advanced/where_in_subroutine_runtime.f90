! vybe-test: fortran/where_advanced/where_in_subroutine_runtime
! origin: languages/fortran/tests/fortran/test_where_advanced.rs

program test
    integer :: a(5) = [3, -1, 5, -2, 4]
    call clamp_negatives(a)
    if ((a(2)) /= 0) then
    print *, "FAIL: want [0] got [", a(2), "]"
    stop 1
end if
    if ((a(4)) /= 0) then
    print *, "FAIL: want [0] got [", a(4), "]"
    stop 1
end if
    if ((a(1)) /= 3) then
    print *, "FAIL: want [3] got [", a(1), "]"
    stop 1
end if
contains
    subroutine clamp_negatives(x)
        integer, intent(inout) :: x(:)
        where (x < 0)
            x = 0
        end where
    end subroutine clamp_negatives
end program test
