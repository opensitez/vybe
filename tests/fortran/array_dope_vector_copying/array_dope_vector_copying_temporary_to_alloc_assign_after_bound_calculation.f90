! vybe-test: fortran/array_dope_vector_copying/array_dope_vector_copying_temporary_to_alloc_assign_after_bound_calculation
! origin: languages/fortran/tests/fortran/test_array_dope_vector_copying.rs

program t
    integer, allocatable :: src(:)
    integer, allocatable :: dst(:)
    integer :: n
    n = 3
    src = (/ 6, 7, 8 /)
    dst = src * n
    if ((size(dst)) /= 3) then
    print *, "FAIL: want [3] got [", size(dst), "]"
    stop 1
end if
    if ((sum(dst)) /= 63) then
    print *, "FAIL: want [63] got [", sum(dst), "]"
    stop 1
end if
    if ((dst(1)) /= 18) then
    print *, "FAIL: want [18] got [", dst(1), "]"
    stop 1
end if
    if ((dst(3)) /= 24) then
    print *, "FAIL: want [24] got [", dst(3), "]"
    stop 1
end if
end program t
