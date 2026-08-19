! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_11_repetition_and_implied_do_mix
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ 2 * 1, (i, i = 2, 4), 3 * 0 /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 11) then
    print *, "FAIL: want [11] got [", sum(values), "]"
    stop 1
end if
    if ((values(2)) /= 2) then
    print *, "FAIL: want [2] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 0) then
    print *, "FAIL: want [0] got [", values(size(values)), "]"
    stop 1
end if
end program t
