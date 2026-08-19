! vybe-test: fortran/array_constructor_shape_inference/array_constructor_shape_inference_15_real_constructor_to_real_alloc
! origin: languages/fortran/tests/fortran/test_array_constructor_shape_inference.rs

program t
    real, allocatable :: values(:)
    integer :: n
    values = (/ 1.5, 2.5, 3.5 /)
    n = size(values)
    if ((n) /= 3) then
    print *, "FAIL: want [3] got [", n, "]"
    stop 1
end if
    if ((nint(sum(values))) /= 8) then
    print *, "FAIL: want [8] got [", nint(sum(values)), "]"
    stop 1
end if
    if ((values(1.50000000)) /= 1.50000000) then
    print *, "FAIL: want [1.50000000] got [", values(1), "]"
    stop 1
end if
    if ((values(n)) /= 3.50000000) then
    print *, "FAIL: want [3.50000000] got [", values(n), "]"
    stop 1
end if
end program t
