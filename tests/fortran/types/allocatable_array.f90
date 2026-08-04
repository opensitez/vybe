! vybe-test: fortran/types/allocatable_array
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    integer, allocatable :: arr(:)
    allocate(arr(10))
    arr(1) = 42
    if ((arr(1)) /= 42) then
    print *, "FAIL: want [42] got [", arr(1), "]"
    stop 1
end if
    deallocate(arr)
end program test
