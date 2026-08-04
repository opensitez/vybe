! vybe-test: fortran/arrays/alloc_1d_runtime_index_write_and_size
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer, allocatable :: v(:)
    allocate(v(3))
    v(1) = 7
    v(2) = 8
    v(3) = 9
    if ((v(1)) /= 7) then
    print *, "FAIL: want [7] got [", v(1), "]"
    stop 1
end if
    if ((v(2)) /= 8) then
    print *, "FAIL: want [8] got [", v(2), "]"
    stop 1
end if
    if ((v(3)) /= 9) then
    print *, "FAIL: want [9] got [", v(3), "]"
    stop 1
end if
    if ((size(v)) /= 3) then
    print *, "FAIL: want [3] got [", size(v), "]"
    stop 1
end if
end program test
