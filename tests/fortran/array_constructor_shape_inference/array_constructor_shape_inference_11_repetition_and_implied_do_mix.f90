! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_11_repetition_and_implied_do_mix
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_11_repetition_and_implied_do_mix
    integer, allocatable :: values(:)
    values = (/ 2 * 1, (i, i = 2, 4), 3 * 0 /)
    if ((size(values)) /= 8) then
    print *, "FAIL: want [8] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 10) then
    print *, "FAIL: want [10] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 0) then
    print *, "FAIL: want [0] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_11_repetition_and_implied_do_mix
