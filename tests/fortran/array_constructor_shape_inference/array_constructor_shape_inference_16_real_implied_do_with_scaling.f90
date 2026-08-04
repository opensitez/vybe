! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_16_real_implied_do_with_scaling
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_16_real_implied_do_with_scaling
    real, allocatable :: values(:)
    integer :: n
    values = (/ (real(i) * 0.5, i = 2, 8, 2) /)
    n = size(values)
    if ((n) /= 4) then
    print *, "FAIL: want [4] got [", n, "]"
    stop 1
end if
    if ((nint(sum(values))) /= 10) then
    print *, "FAIL: want [10] got [", nint(sum(values)), "]"
    stop 1
end if
    if ((values(1)) /= 1) then
    print *, "FAIL: want [1] got [", values(1), "]"
    stop 1
end if
    if ((values(n)) /= 4) then
    print *, "FAIL: want [4] got [", values(n), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_16_real_implied_do_with_scaling
