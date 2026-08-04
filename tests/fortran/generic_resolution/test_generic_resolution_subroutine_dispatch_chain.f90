! vybe-test: fortran/generic_resolution/test_generic_resolution_subroutine_dispatch_chain
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs

module m
    interface g
        module procedure ss, sr
    end interface
contains
    subroutine ss(n)
        integer, intent(in) :: n
        if ((n * 2) /= 10) then
    print *, "FAIL: want [10] got [", n * 2, "]"
    stop 1
end if
    end subroutine

    subroutine sr(r)
        real, intent(in) :: r
        if ((nint(r * 2.0)) /= 12) then
    print *, "FAIL: want [12] got [", nint(r * 2.0), "]"
    stop 1
end if
    end subroutine
end module m

program test_generic_resolution_subroutine_dispatch_chain
    use m
    call g(5)
    call g(6.0)
end program test_generic_resolution_subroutine_dispatch_chain
