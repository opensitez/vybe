! vybe-test: fortran/generic_interfaces/test_generic_interfaces_dispatched_subroutines
! origin: languages/fortran/tests/fortran/test_generic_interfaces.rs

module m
    interface g
        module procedure si, sr
    end interface
contains
    subroutine si(v)
        integer, intent(in) :: v
        if ((v + 1) /= 4) then
    print *, "FAIL: want [4] got [", v + 1, "]"
    stop 1
end if
    end subroutine

    subroutine sr(v)
        real, intent(in) :: v
        if ((nint(v) + 1) /= 5) then
    print *, "FAIL: want [5] got [", nint(v) + 1, "]"
    stop 1
end if
    end subroutine
end module m

program test_generic_interfaces_dispatched_subroutines
    use m
    call g(3)
    call g(4.0)
end program test_generic_interfaces_dispatched_subroutines
