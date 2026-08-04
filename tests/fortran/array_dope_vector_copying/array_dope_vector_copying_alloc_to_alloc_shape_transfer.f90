! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_alloc_to_alloc_shape_transfer
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_alloc_to_alloc_shape_transfer
    integer, allocatable :: source(:), target(:)
    source = (/ 4, 8, 12, 16 /)
    target = source
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 40) then
    print *, "FAIL: want [40] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 4) then
    print *, "FAIL: want [4] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 16) then
    print *, "FAIL: want [16] got [", target(size(target)), "]"
    stop 1
end if
end program array_dope_vector_copying_alloc_to_alloc_shape_transfer
