! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_empty_section_to_allocatable_target
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_empty_section_to_allocatable_target
    integer, allocatable :: source(:), target(:)
    source = (/ 9, 8, 7 /)
    target = source(3:2)
    if ((size(target)) /= 0) then
    print *, "FAIL: want [0] got [", size(target), "]"
    stop 1
end if
    if ((merge(1, 0, size(target) == 0)) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, size(target) == 0), "]"
    stop 1
end if
end program array_dope_vector_copying_empty_section_to_allocatable_target
