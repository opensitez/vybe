! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_08_implied_do_square_series
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_08_implied_do_square_series
    integer, allocatable :: values(:)
    values = (/ (i * i, i = 1, 4) /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 30) then
    print *, "FAIL: want [30] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 16) then
    print *, "FAIL: want [16] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_08_implied_do_square_series
