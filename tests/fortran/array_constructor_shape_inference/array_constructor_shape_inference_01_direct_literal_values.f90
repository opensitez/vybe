! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_01_direct_literal_values
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_01_direct_literal_values
    integer, allocatable :: values(:)
    values = (/ 1, 2, 3 /)
    if ((size(values)) /= 3) then
    print *, "FAIL: want [3] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 6) then
    print *, "FAIL: want [6] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 3) then
    print *, "FAIL: want [3] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_01_direct_literal_values
