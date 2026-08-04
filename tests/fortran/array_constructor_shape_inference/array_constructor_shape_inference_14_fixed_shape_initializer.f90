! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_14_fixed_shape_initializer
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_14_fixed_shape_initializer
    integer :: values(4) = (/ 2, 4, 6, 8 /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 20) then
    print *, "FAIL: want [20] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 8) then
    print *, "FAIL: want [8] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_14_fixed_shape_initializer
