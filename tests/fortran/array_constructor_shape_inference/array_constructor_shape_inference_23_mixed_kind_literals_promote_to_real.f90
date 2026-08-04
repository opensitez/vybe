! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_23_mixed_kind_literals_promote_to_real
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program test_array_constructor_shape_inference_23_mixed_kind_literals_promote_to_real
    real, allocatable :: values(:)
    values = (/ 1, 2.25, 4 /)
    if ((size(values)) /= 3) then
    print *, "FAIL: want [3] got [", size(values), "]"
    stop 1
end if
    if ((nint(sum(values))) /= 7) then
    print *, "FAIL: want [7] got [", nint(sum(values)), "]"
    stop 1
end if
    if ((ceiling(values(2))) /= 3) then
    print *, "FAIL: want [3] got [", ceiling(values(2)), "]"
    stop 1
end if
    if ((floor(values(1))) /= 1) then
    print *, "FAIL: want [1] got [", floor(values(1)), "]"
    stop 1
end if
end program test_array_constructor_shape_inference_23_mixed_kind_literals_promote_to_real
