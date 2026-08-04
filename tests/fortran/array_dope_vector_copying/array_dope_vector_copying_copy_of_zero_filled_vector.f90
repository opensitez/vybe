! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_of_zero_filled_vector
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_of_zero_filled_vector
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 0, 0, 0 /)
    target = source
    if ((size(target)) /= 3) then
    print *, "FAIL: want [3] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 0) then
    print *, "FAIL: want [0] got [", sum(target), "]"
    stop 1
end if
    if ((target(1) + target(size(target))) /= 0) then
    print *, "FAIL: want [0] got [", target(1) + target(size(target)), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_of_zero_filled_vector
