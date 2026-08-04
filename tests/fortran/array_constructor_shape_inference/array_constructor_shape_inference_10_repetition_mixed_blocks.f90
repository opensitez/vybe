! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_10_repetition_mixed_blocks
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_10_repetition_mixed_blocks
    integer, allocatable :: values(:)
    values = (/ 2 * 3, 3 * 8 /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 34) then
    print *, "FAIL: want [34] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 3) then
    print *, "FAIL: want [3] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 8) then
    print *, "FAIL: want [8] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_10_repetition_mixed_blocks
