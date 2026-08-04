! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_section_to_mismatched_lbound
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_section_to_mismatched_lbound
    integer, allocatable :: source(:)
    integer :: target(10:12)
    source = (/ 7, 8, 9 /)
    target = source(2:4)
    if ((source(1)) /= 7) then
    print *, "FAIL: want [7] got [", source(1), "]"
    stop 1
end if
    if ((source(3)) /= 9) then
    print *, "FAIL: want [9] got [", source(3), "]"
    stop 1
end if
    if ((target(10)) /= 8) then
    print *, "FAIL: want [8] got [", target(10), "]"
    stop 1
end if
    if ((target(12)) /= 9) then
    print *, "FAIL: want [9] got [", target(12), "]"
    stop 1
end if
end program array_dope_vector_copying_section_to_mismatched_lbound
