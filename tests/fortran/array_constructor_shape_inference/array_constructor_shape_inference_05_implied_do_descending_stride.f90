! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_05_implied_do_descending_stride
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_05_implied_do_descending_stride
    integer, allocatable :: values(:)
    values = (/ (i, i = 11, 3, -3) /)
    if ((size(values)) /= 3) then
    print *, "FAIL: want [3] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 17) then
    print *, "FAIL: want [17] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 11) then
    print *, "FAIL: want [11] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 5) then
    print *, "FAIL: want [5] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_05_implied_do_descending_stride
