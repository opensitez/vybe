! vybe-test: fortran/derived_types_advanced/allocatable_type_array
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Node
        integer :: value
    end type Node
    type(Node), allocatable :: nodes(:)
    allocate(nodes(5))
    nodes(1)%value = 42
    if ((nodes(1)%value) /= 42) then
    print *, "FAIL: want [42] got [", nodes(1)%value, "]"
    stop 1
end if
    nodes(5)%value = 9
    if ((nodes(5)%value) /= 9) then
    print *, "FAIL: want [9] got [", nodes(5)%value, "]"
    stop 1
end if
    deallocate(nodes)
end program test
