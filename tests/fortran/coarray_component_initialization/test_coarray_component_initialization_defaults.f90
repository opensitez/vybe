! vybe-test: fortran/coarray_component_initialization/test_coarray_component_initialization_defaults
! origin: languages/fortran/tests/fortran/test_coarray_component_initialization.rs

program test_coarray_component_initialization
    type :: endpoint
        integer :: value
        integer, allocatable :: values(:)
    end type endpoint

    type(endpoint) :: x
    allocate(x%values(3))
    x%values = (/1, 2, 3/)
    x%value = x%values(2)

    if ((x%value) /= 2) then
    print *, "FAIL: want [2] got [", x%value, "]"
    stop 1
end if
    if ((x%values(3)) /= 3) then
    print *, "FAIL: want [3] got [", x%values(3), "]"
    stop 1
end if
end program test_coarray_component_initialization
