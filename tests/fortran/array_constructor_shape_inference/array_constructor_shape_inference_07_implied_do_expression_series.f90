! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_07_implied_do_expression_series
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ (i * 3 - 1, i = 1, 5) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 40) then
    print *, "FAIL: want [40] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 14) then
    print *, "FAIL: want [14] got [", values(size(values)), "]"
    stop 1
end if
end program t
