! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_20_character_vector_shape_inference
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_20_character_vector_shape_inference
    character(len=4), allocatable :: values(:)
    values = (/ 'ab', 'cde', 'x' /)
    if ((size(values)) /= 3) then
    print *, "FAIL: want [3] got [", size(values), "]"
    stop 1
end if
    if ((len(values(1))) /= 4) then
    print *, "FAIL: want [4] got [", len(values(1)), "]"
    stop 1
end if
    if ((len_trim(values(2))) /= 3) then
    print *, "FAIL: want [3] got [", len_trim(values(2)), "]"
    stop 1
end if
    if ((len_trim(values(3))) /= 1) then
    print *, "FAIL: want [1] got [", len_trim(values(3)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_20_character_vector_shape_inference
