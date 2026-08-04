! vybe-test: fortran/array_bounds_and_lbound_ubound/array_bounds_and_lbound_ubound_result_vector_without_dim
! origin: languages/fortran/tests/fortran/test_array_bounds_and_lbound_ubound.rs

program array_bounds_and_lbound_ubound_result_vector_without_dim
    integer :: a(-3:0, 10:12)
    integer :: lb(2), ub(2)
    lb = lbound(a)
    ub = ubound(a)
    if ((size(lb)) /= 2) then
    print *, "FAIL: want [2] got [", size(lb), "]"
    stop 1
end if
    if ((lb(1)) /= -3) then
    print *, "FAIL: want [-3] got [", lb(1), "]"
    stop 1
end if
    if ((lb(2)) /= 10) then
    print *, "FAIL: want [10] got [", lb(2), "]"
    stop 1
end if
    if ((size(ub)) /= 2) then
    print *, "FAIL: want [2] got [", size(ub), "]"
    stop 1
end if
    if ((ub(1)) /= 0) then
    print *, "FAIL: want [0] got [", ub(1), "]"
    stop 1
end if
    if ((ub(2)) /= 12) then
    print *, "FAIL: want [12] got [", ub(2), "]"
    stop 1
end if
end program array_bounds_and_lbound_ubound_result_vector_without_dim
