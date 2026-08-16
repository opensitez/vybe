! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_03_implied_do_linear_sequence
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ (i, i = 1, 5) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 15) then
    print *, "FAIL: want [15] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 5) then
    print *, "FAIL: want [5] got [", values(size(values)), "]"
    stop 1
end if
end program t
