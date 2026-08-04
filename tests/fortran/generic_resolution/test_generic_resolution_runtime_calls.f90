! vybe-test: fortran/generic_resolution/test_generic_resolution_runtime_calls
! origin: languages/fortran/tests/fortran/test_generic_resolution.rs

module m
    interface g
        module procedure si, sr
    end interface
contains
    integer function si(i)
        integer, intent(in) :: i
        si = i + 10
    end function

    real function sr(r)
        real, intent(in) :: r
        sr = r + 1.0
    end function
end module m

program test_generic_resolution_runtime_calls
    use m
    if ((g(1)) /= 11) then
    print *, "FAIL: want [11] got [", g(1), "]"
    stop 1
end if
    if ((nint(g(3.0))) /= 4) then
    print *, "FAIL: want [4] got [", nint(g(3.0)), "]"
    stop 1
end if
end program test_generic_resolution_runtime_calls
