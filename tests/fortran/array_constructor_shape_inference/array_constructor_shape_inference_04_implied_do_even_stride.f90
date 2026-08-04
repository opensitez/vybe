! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_04_implied_do_even_stride
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_04_implied_do_even_stride
    integer, allocatable :: values(:)
    values = (/ (i, i = 2, 10, 2) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 30) then
    print *, "FAIL: want [30] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 10) then
    print *, "FAIL: want [10] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_04_implied_do_even_stride
