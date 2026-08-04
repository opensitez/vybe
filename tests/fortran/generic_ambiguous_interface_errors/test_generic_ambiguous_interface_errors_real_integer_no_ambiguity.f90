! vybe-test: fortran/generic_ambiguous_interface_errors/test_generic_ambiguous_interface_errors_real_integer_no_ambiguity
! origin: languages/fortran/tests/fortran/test_generic_ambiguous_interface_errors.rs

program test_generic_ambiguous_interface_errors
    if ((magnitude(2)) /= 2) then
    print *, "FAIL: want [2] got [", magnitude(2), "]"
    stop 1
end if
    if ((nint(magnitude(-3.0))) /= 3) then
    print *, "FAIL: want [3] got [", nint(magnitude(-3.0)), "]"
    stop 1
end if

contains
    interface magnitude
        module procedure abs_int
        module procedure abs_real
    end interface

    integer function abs_int(v)
        integer, intent(in) :: v
        abs_int = abs(v)
    end function

    real function abs_real(v)
        real, intent(in) :: v
        abs_real = abs(v)
    end function
end program test_generic_ambiguous_interface_errors
