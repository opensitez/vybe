! vybe-test: fortran/arrays/top_level_deallocate_nulls_allocatable_array
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer, allocatable :: arr(:)
    allocate(arr(2))
    arr(1) = 5
    deallocate(arr)
    if ((allocated(arr)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", allocated(arr), "]"
    stop 1
end if
end program test
