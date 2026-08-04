! vybe-test: fortran/generic_ambiguity/generic_ambiguity_function_call_runtime_dispatch
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs

module m
    interface g
        module procedure gi, gr
    end interface
contains
    integer function gi(i)
        integer, intent(in) :: i
        gi = i + 10
    end function

    real function gr(r)
        real, intent(in) :: r
        gr = r + 1.5
    end function
end module m

program test_generic_ambiguity_function_call_runtime_dispatch
    use m
    if ((g(4)) /= 14) then
    print *, "FAIL: want [14] got [", g(4), "]"
    stop 1
end if
    if ((nint(g(2.0))) /= 4) then
    print *, "FAIL: want [4] got [", nint(g(2.0)), "]"
    stop 1
end if
end program test_generic_ambiguity_function_call_runtime_dispatch
