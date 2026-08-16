! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_22_descending_implied_do_with_variables
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer :: start_idx
    integer :: stop_idx
    integer :: step
    integer, allocatable :: values(:)
    start_idx = 9
    stop_idx = 1
    step = -2
    values = (/ (i, i = start_idx, stop_idx, step) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 25) then
    print *, "FAIL: want [25] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 9) then
    print *, "FAIL: want [9] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 1) then
    print *, "FAIL: want [1] got [", values(size(values)), "]"
    stop 1
end if
end program t
