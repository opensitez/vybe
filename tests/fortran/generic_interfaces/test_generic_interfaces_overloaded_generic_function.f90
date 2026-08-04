! vybe-test: fortran/generic_interfaces/test_generic_interfaces_overloaded_generic_function
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs

module m
    interface val
        module procedure vi
        module procedure vr
    end interface
contains
    integer function vi(v)
        integer, intent(in) :: v
        vi = v * 2
    end function

    real function vr(v)
        real, intent(in) :: v
        vr = v * 2.0
    end function
end module m

program test_generic_interfaces_overloaded_generic_function
    use m
    if ((val(4)) /= 8) then
    print *, "FAIL: want [8] got [", val(4), "]"
    stop 1
end if
    if ((nint(val(2.5))) /= 5) then
    print *, "FAIL: want [5] got [", nint(val(2.5)), "]"
    stop 1
end if
end program test_generic_interfaces_overloaded_generic_function
