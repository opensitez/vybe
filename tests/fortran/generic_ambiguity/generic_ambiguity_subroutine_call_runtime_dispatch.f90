! vybe-test: fortran/generic_ambiguity/generic_ambiguity_subroutine_call_runtime_dispatch
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs

module m
    interface g
        module procedure si, sr
    end interface
contains
    subroutine si(i)
        integer, intent(in) :: i
        if ((i + 10) /= 12) then
    print *, "FAIL: want [12] got [", i + 10, "]"
    stop 1
end if
    end subroutine

    subroutine sr(r)
        real, intent(in) :: r
        if ((nint(r) + 20) /= 23) then
    print *, "FAIL: want [23] got [", nint(r) + 20, "]"
    stop 1
end if
    end subroutine
end module m

program test_generic_ambiguity_subroutine_call_runtime_dispatch
    use m
    call g(2)
    call g(3.0)
end program test_generic_ambiguity_subroutine_call_runtime_dispatch
