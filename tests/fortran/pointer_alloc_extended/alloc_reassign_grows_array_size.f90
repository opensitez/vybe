! vybe-test: fortran/pointer_alloc_extended/alloc_reassign_grows_array_size
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, allocatable :: items(:)
items = [1, 2]
if ((size(items)) /= 2) then
    print *, "FAIL: want [2] got [", size(items), "]"
    stop 1
end if
items = [10, 20, 30, 40]
if ((size(items)) /= 4) then
    print *, "FAIL: want [4] got [", size(items), "]"
    stop 1
end if
if ((items(4)) /= 40) then
    print *, "FAIL: want [40] got [", items(4), "]"
    stop 1
end if
end program t
