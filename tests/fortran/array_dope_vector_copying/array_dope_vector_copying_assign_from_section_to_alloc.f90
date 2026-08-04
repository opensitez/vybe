! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_assign_from_section_to_alloc
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_assign_from_section_to_alloc
    integer, allocatable :: source(:)
    integer, allocatable :: target(:)
    source = (/ 10, 20, 30, 40, 50, 60 /)
    target = source(2:5)
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 140) then
    print *, "FAIL: want [140] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 20) then
    print *, "FAIL: want [20] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 50) then
    print *, "FAIL: want [50] got [", target(size(target)), "]"
    stop 1
end if
end program array_dope_vector_copying_assign_from_section_to_alloc
