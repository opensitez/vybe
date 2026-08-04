! vybe-test: fortran/pointers/allocatable_in_type
! origin: languages/fortran/tests/fortran/test_pointers.rs

program test
    type :: DynList
        integer, allocatable :: items(:)
        integer :: count = 0
    end type DynList
    type(DynList) :: list
    allocate(list%items(10))
    list%items(1) = 100
    list%count = 1
    if ((list%items(1)) /= 100) then
    print *, "FAIL: want [100] got [", list%items(1), "]"
    stop 1
end if
    if ((list%count) /= 1) then
    print *, "FAIL: want [1] got [", list%count, "]"
    stop 1
end if
    deallocate(list%items)
end program test
