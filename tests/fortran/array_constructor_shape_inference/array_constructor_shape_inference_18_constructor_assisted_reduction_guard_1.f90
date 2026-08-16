! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_18_constructor_assisted_reduction_guard_1
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    call validate_values((/ 10, 20, 30, 40, 50 /), 50, 150)
contains
    subroutine validate_values(values, expected_last, expected_sum)
        integer, intent(in) :: values(:)
        integer, intent(in) :: expected_last
        integer, intent(in) :: expected_sum
        integer :: last_value
        last_value = values(size(values))
        if ((merge(1, 0, size(values) >= 3)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, size(values) >= 3), "]"
    stop 1
end if
        if ((merge(1, 0, sum(values) == expected_sum)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, sum(values) == expected_sum), "]"
    stop 1
end if
        if ((merge(1, 0, last_value == expected_last)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, last_value == expected_last), "]"
    stop 1
end if
    end subroutine validate_values
end program t
