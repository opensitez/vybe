! vybe-test: fortran/generic_interface_resolution/test_generic_interface_resolution_selects_real_and_integer_bindings
! origin: languages/fortran/tests/fortran/test_generic_interface_resolution.rs

program test_generic_interface_resolution_selects_real_and_integer_bindings
    if ((add(2, 3)) /= 5) then
    print *, "FAIL: want [5] got [", add(2, 3), "]"
    stop 1
end if
    if ((nint(add(2.0, 3.5))) /= 6) then
    print *, "FAIL: want [6] got [", nint(add(2.0, 3.5)), "]"
    stop 1
end if

contains
    interface add
        module procedure add_int
        module procedure add_real
    end interface

    integer function add_int(a, b)
        integer, intent(in) :: a, b
        add_int = a + b
    end function

    real function add_real(a, b)
        real, intent(in) :: a, b
        add_real = a + b
    end function
end program test_generic_interface_resolution_selects_real_and_integer_bindings
