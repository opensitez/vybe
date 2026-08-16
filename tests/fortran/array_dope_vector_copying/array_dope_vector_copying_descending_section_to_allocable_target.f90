! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_descending_section_to_allocable_target
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer, allocatable :: source(:), target(:)
    source = (/ 1, 2, 3, 4, 5, 6 /)
    target = source(6:3:-1)
    if ((size(target)) /= 4) then
    print *, "FAIL: want [4] got [", size(target), "]"
    stop 1
end if
    if ((sum(target)) /= 18) then
    print *, "FAIL: want [18] got [", sum(target), "]"
    stop 1
end if
    if ((target(1)) /= 6) then
    print *, "FAIL: want [6] got [", target(1), "]"
    stop 1
end if
    if ((target(size(target))) /= 3) then
    print *, "FAIL: want [3] got [", target(size(target)), "]"
    stop 1
end if
end program t
