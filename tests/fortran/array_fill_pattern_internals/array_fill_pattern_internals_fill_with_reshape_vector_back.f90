! vybe-test: fortran/array_fill_pattern_internals/array_fill_pattern_internals_fill_with_reshape_vector_back
! origin: languages/fortran/tests/fortran/test_array_fill_pattern_internals.rs

program array_fill_pattern_internals_fill_with_reshape_vector_back
    integer, allocatable :: flat(:)
    integer :: mat(2,3)
    flat = (/ 1, 2, 3, 4, 5, 6 /)
    mat = reshape(flat, (/2,3/))
    flat = mat
    if ((sum(flat)) /= 21) then
    print *, "FAIL: want [21] got [", sum(flat), "]"
    stop 1
end if
    if ((flat(1)) /= 1) then
    print *, "FAIL: want [1] got [", flat(1), "]"
    stop 1
end if
    if ((flat(6)) /= 6) then
    print *, "FAIL: want [6] got [", flat(6), "]"
    stop 1
end if
end program array_fill_pattern_internals_fill_with_reshape_vector_back
