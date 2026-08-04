! vybe-test: fortran/derived_type_pointer_components/test_derived_type_pointer_components_reference_and_read
! origin: languages/fortran/tests/fortran/test_derived_type_pointer_components.rs

program test_derived_type_pointer_components
    type :: container
        integer, pointer :: values(:)
    end type

    integer, target :: storage(3)
    type(container) :: box

    storage = (/10, 20, 30/)
    box%values => storage
    if ((box%values(2)) /= 20) then
    print *, "FAIL: want [20] got [", box%values(2), "]"
    stop 1
end if
end program test_derived_type_pointer_components
