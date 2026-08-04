! vybe-test: fortran/generic_interface_resolution/test_generic_interface_resolution_selects_integer_binding
! origin: languages/fortran/tests/fortran/test_generic_interface_resolution.rs

program test_generic_interface_resolution
    integer :: result
    result = add(2, 3)
    if ((result) /= 5) then
    print *, "FAIL: want [5] got [", result, "]"
    stop 1
end if

contains
    interface add
        module procedure add_int
    end interface

    integer function add_int(a, b)
        integer, intent(in) :: a, b
        add_int = a + b
    end function
end program test_generic_interface_resolution
