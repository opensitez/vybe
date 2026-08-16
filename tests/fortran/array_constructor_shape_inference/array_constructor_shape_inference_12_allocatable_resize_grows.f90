! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_12_allocatable_resize_grows
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    integer, allocatable :: values(:)
    values = (/ 11, 22 /)
    if ((size(values)) /= 2) then
    print *, "FAIL: want [2] got [", size(values), "]"
    stop 1
end if
    values = (/ 1, 2, 3, 4, 5, 6 /)
    if ((size(values)) /= 6) then
    print *, "FAIL: want [6] got [", size(values), "]"
    stop 1
end if
    if ((sum(values)) /= 21) then
    print *, "FAIL: want [21] got [", sum(values), "]"
    stop 1
end if
    if ((values(size(values))) /= 6) then
    print *, "FAIL: want [6] got [", values(size(values)), "]"
    stop 1
end if
end program t
