! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_09_repetition_single_block
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ 4 * 7 /)
    if ((size(values)) /= 1) then
    print *, "FAIL: want [1] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 28) then
    print *, "FAIL: want [28] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 28) then
    print *, "FAIL: want [28] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 28) then
    print *, "FAIL: want [28] got [", values(size(values)), "]"
    stop 1
end if
end program t
