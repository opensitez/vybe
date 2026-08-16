! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_17_constructor_as_subroutine_argument_size_and_sum
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    call validate_values((/ 3, 1, 4, 1, 5, 9 /), 6, 23, 3)
contains
    subroutine validate_values(values, expected_size, expected_sum, expected_first)
        integer, intent(in) :: values(:)
        integer, intent(in) :: expected_size
        integer, intent(in) :: expected_sum
        integer, intent(in) :: expected_first
        if ((merge(1, 0, size(values) == expected_size)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, size(values) == expected_size), "]"
    stop 1
end if
        if ((merge(1, 0, sum(values) == expected_sum)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, sum(values) == expected_sum), "]"
    stop 1
end if
        if ((merge(1, 0, values(1) == expected_first)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, values(1) == expected_first), "]"
    stop 1
end if
    end subroutine validate_values
end program t
