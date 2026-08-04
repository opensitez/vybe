! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_before_and_after_realloc
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_before_and_after_realloc
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 1, 2, 3 /)
    target = source
    source = (/ 8, 9, 10, 11 /)
    if ((sum(target)) /= 6) then
    print *, "FAIL: want [6] got [", sum(target), "]"
    stop 1
end if
    if ((size(source)) /= 4) then
    print *, "FAIL: want [4] got [", size(source), "]"
    stop 1
end if
    if ((size(target)) /= 3) then
    print *, "FAIL: want [3] got [", size(target), "]"
    stop 1
end if
    if ((target(2)) /= 2) then
    print *, "FAIL: want [2] got [", target(2), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_before_and_after_realloc
