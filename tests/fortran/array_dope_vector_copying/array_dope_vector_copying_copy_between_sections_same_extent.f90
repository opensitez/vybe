! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_copy_between_sections_same_extent
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program array_dope_vector_copying_copy_between_sections_same_extent
    integer :: source(1:8)
    integer :: work(0:7)
    source = (/ 1, 2, 3, 4, 5, 6, 7, 8 /)
    work(2:6) = source(3:7)
    if ((sum(work)) /= 22) then
    print *, "FAIL: want [22] got [", sum(work), "]"
    stop 1
end if
    if ((work(2)) /= 3) then
    print *, "FAIL: want [3] got [", work(2), "]"
    stop 1
end if
    if ((work(6)) /= 7) then
    print *, "FAIL: want [7] got [", work(6), "]"
    stop 1
end if
end program array_dope_vector_copying_copy_between_sections_same_extent
