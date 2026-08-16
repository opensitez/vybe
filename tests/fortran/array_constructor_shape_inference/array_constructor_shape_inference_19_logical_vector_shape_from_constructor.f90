! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_19_logical_vector_shape_from_constructor
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    logical, allocatable :: values(:)
    values = (/ .true., .false., .true., .true., .false. /)
    if ((size(values)) /= 5) then
    print *, "FAIL: want [5] got [", size(values), "]"
    stop 1
end if
    if ((count(values)) /= 3) then
    print *, "FAIL: want [3] got [", count(values), "]"
    stop 1
end if
    if ((merge(1, 0, values(1))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, values(1)), "]"
    stop 1
end if
    if ((merge(1, 0, values(5))) /= 0) then
    print *, "FAIL: want [0] got [", merge(1, 0, values(5)), "]"
    stop 1
end if
end program t
