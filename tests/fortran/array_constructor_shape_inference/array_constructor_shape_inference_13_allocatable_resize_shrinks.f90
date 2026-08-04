! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_13_allocatable_resize_shrinks
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_13_allocatable_resize_shrinks
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3, 4, 5, 6 /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    values = (/ 99 /)
    if ((size(values)) /= 1) then
    print *, "FAIL: want [1] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 99) then
    print *, "FAIL: want [99] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 99) then
    print *, "FAIL: want [99] got [", values(1), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_13_allocatable_resize_shrinks
