! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_10_repetition_mixed_blocks
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ 2 * 3, 3 * 8 /)
    if ((size(values)) /= 2) then
    print *, "FAIL: want [2] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 30) then
    print *, "FAIL: want [30] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 6) then
    print *, "FAIL: want [6] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 24) then
    print *, "FAIL: want [24] got [", values(size(values)), "]"
    stop 1
end if
end program t
