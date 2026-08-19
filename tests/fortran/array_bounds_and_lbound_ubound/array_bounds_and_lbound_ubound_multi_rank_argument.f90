! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_multi_rank_argument
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_multi_rank_argument
    integer :: data(-2:1, 6:9, 0:0)
    call dump_bounds(data)

contains
    subroutine dump_bounds(x)
        integer, intent(in) :: x(:, :, :)
        if ((lbound(x, 1)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(x, 1), "]"
    stop 1
end if
        if ((ubound(x, 1)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(x, 1), "]"
    stop 1
end if
        if ((lbound(x, 2)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(x, 2), "]"
    stop 1
end if
        if ((ubound(x, 2)) /= 4) then
    print *, "FAIL: want [4] got [", ubound(x, 2), "]"
    stop 1
end if
        if ((lbound(x, 3)) /= 1) then
    print *, "FAIL: want [1] got [", lbound(x, 3), "]"
    stop 1
end if
        if ((ubound(x, 3)) /= 1) then
    print *, "FAIL: want [1] got [", ubound(x, 3), "]"
    stop 1
end if
    end subroutine dump_bounds
end program array_bounds_and_lbound_ubound_multi_rank_argument
