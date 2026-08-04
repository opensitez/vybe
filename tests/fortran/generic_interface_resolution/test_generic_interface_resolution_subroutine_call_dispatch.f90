! vybe-test: fortran/generic_interface_resolution/test_generic_interface_resolution_subroutine_call_dispatch
! origin: languages/fortran/tests/fortran/test_generic_interface_resolution.rs

program test_generic_interface_resolution_subroutine_call_dispatch
    call scale_out(2)
    call scale_out(3.0)

contains
    interface scale_out
        module procedure scale_i
        module procedure scale_r
    end interface

    subroutine scale_i(value)
        integer, intent(in) :: value
        if ((value * 10) /= 20) then
    print *, "FAIL: want [20] got [", value * 10, "]"
    stop 1
end if
    end subroutine

    subroutine scale_r(value)
        real, intent(in) :: value
        if ((nint(value * 10.0)) /= 30) then
    print *, "FAIL: want [30] got [", nint(value * 10.0), "]"
    stop 1
end if
    end subroutine
end program test_generic_interface_resolution_subroutine_call_dispatch
