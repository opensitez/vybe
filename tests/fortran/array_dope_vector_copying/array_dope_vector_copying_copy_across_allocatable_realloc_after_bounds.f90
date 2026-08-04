! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    integer :: sum_target
    source = (/ 3, 1, 4, 1, 5, 9 /)
    target = source(2:5)
    target = source
    sum_target = sum(target)
    if ((size(target)) /= 6) then
    print *, "FAIL: want [6] got [", size(target), "]"
    stop 1
end if
    if ((sum_target) /= 23) then
    print *, "FAIL: want [23] got [", sum_target, "]"
    stop 1
end if
    if ((target(1)) /= 3) then
    print *, "FAIL: want [3] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 9) then
    print *, "FAIL: want [9] got [", target(size(target)), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_across_allocatable_realloc_after_bounds
