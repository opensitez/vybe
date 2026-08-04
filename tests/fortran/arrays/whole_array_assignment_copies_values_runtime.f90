! vybe-test: fortran/arrays/whole_array_assignment_copies_values_runtime
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    integer :: source(4) = [1, 2, 3, 4]
    integer, allocatable :: copy(:)

    allocate(copy(4))
    copy = source
    copy(2) = 99

    if ((source(2)) /= 2) then
    print *, "FAIL: want [2] got [", source(2), "]"
    stop 1
end if
    if ((copy(2)) /= 99) then
    print *, "FAIL: want [99] got [", copy(2), "]"
    stop 1
end if
end program test
