! vybe-test: fortran/derived_type_memory_layout/test_derived_type_memory_layout_reports_component_storage
! origin: languages/fortran/tests/fortran/test_derived_type_memory_layout.rs

program test_derived_type_memory_layout
    type :: item
        integer :: a
        real :: b
    end type

    type(item) :: v
    if ((storage_size(v)) /= 64) then
    print *, "FAIL: want [64] got [", storage_size(v), "]"
    stop 1
end if
    if ((storage_size(v%a)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(v%a), "]"
    stop 1
end if
    if ((storage_size(v%b)) /= 32) then
    print *, "FAIL: want [32] got [", storage_size(v%b), "]"
    stop 1
end if
end program test_derived_type_memory_layout
