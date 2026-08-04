! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_via_temporary_expression
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_via_temporary_expression
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 5, 10, 15, 20 /)
    target = source + 1
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 54) then
    print *, "FAIL: want [54] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 6) then
    print *, "FAIL: want [6] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 21) then
    print *, "FAIL: want [21] got [", target(size(target)), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_via_temporary_expression
