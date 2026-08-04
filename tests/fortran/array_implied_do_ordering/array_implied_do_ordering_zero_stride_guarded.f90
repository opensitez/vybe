! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_zero_stride_guarded
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_zero_stride_guarded
    integer :: status
    integer :: values(3)
    if (1 <= 3) then
        values = [(1, i = 1, 3)]
    else
        values = 0
    end if
    status = sum(values)
    if ((status) /= 3) then
    print *, "FAIL: want [3] got [", status, "]"
    stop 1
end if
    if ((values(3)) /= 1) then
    print *, "FAIL: want [1] got [", values(3), "]"
    stop 1
end if
end program array_implied_do_ordering_zero_stride_guarded
