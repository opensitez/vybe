! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_02_direct_literal_with_negatives
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_02_direct_literal_with_negatives
    integer, allocatable :: values(:)
    values = (/ -3, 5, -1, 7 /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 8) then
    print *, "FAIL: want [8] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= -3) then
    print *, "FAIL: want [-3] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 7) then
    print *, "FAIL: want [7] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_02_direct_literal_with_negatives
