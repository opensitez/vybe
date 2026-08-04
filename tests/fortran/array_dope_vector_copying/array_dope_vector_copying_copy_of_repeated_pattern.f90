! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_of_repeated_pattern
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_of_repeated_pattern
    integer :: source(1:4)
    integer :: target(1:4)
    source = (/ 2 * 1, 2 * 4 /)
    target = source
    if ((merge(1, 0, all(target == source))) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, all(target == source)), "]"
    stop 1
end if
    if ((target(1) + target(4)) /= 5) then
    print *, "FAIL: want [5] got [", target(1) + target(4), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_of_repeated_pattern
