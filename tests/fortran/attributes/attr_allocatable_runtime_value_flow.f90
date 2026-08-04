! vybe-test: fortran/attributes/attr_allocatable_runtime_value_flow
! origin: languages/fortran/tests/fortran/test_attributes.rs

program attr_allocatable_runtime_value_flow
    integer, allocatable :: values(:)
    allocate(values(2))
    values = (/10, 20/)
    if ((values(1)) /= 10) then
    print *, "FAIL: want [10] got [", values(1), "]"
    stop 1
end if
    if ((values(2)) /= 20) then
    print *, "FAIL: want [20] got [", values(2), "]"
    stop 1
end if
end program attr_allocatable_runtime_value_flow
