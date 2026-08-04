! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_09_repetition_single_block
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_09_repetition_single_block
    integer, allocatable :: values(:)
    values = (/ 4 * 7 /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 28) then
    print *, "FAIL: want [28] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 7) then
    print *, "FAIL: want [7] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 7) then
    print *, "FAIL: want [7] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_09_repetition_single_block
