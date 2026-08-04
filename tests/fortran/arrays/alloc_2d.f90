! vybe-test: fortran/arrays/alloc_2d
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer, allocatable :: m(:,:)
    allocate(m(3,3))
    m(1,1) = 7
    m(3,3) = 9
    if ((m(1,1)) /= 7) then
    print *, "FAIL: want [7] got [", m(1,1), "]"
    stop 1
end if
    if ((m(3,3)) /= 9) then
    print *, "FAIL: want [9] got [", m(3,3), "]"
    stop 1
end if
    deallocate(m)
end program test
