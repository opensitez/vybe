! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_06_implied_do_from_variables
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_06_implied_do_from_variables
    integer :: start_idx
    integer :: stop_idx
    integer :: step
    integer, allocatable :: values(:)
    start_idx = 4
    stop_idx = 12
    step = 2
    values = (/ (i, i = start_idx, stop_idx, step) /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 40) then
    print *, "FAIL: want [40] got [", sum(values), "]"
    stop 1
end if
    if ((values(1)) /= 4) then
    print *, "FAIL: want [4] got [", values(1), "]"
    stop 1
end if
    if ((values(size(values))) /= 12) then
    print *, "FAIL: want [12] got [", values(size(values)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_06_implied_do_from_variables
