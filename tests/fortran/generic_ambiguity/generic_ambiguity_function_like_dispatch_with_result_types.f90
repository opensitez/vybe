! vybe-test: fortran/generic_ambiguity/generic_ambiguity_function_like_dispatch_with_result_types
! origin: languages/fortran/tests/fortran/test_generic_ambiguity.rs

module m
    interface g
        module procedure i2, r2
    end interface
contains
    integer function i2(i)
        integer, intent(in) :: i
        i2 = i * 2
    end function

    real function r2(r)
        real, intent(in) :: r
        r2 = r * 3.0
    end function
end module m

program test_generic_ambiguity_function_like_dispatch_with_result_types
    use m
    if ((g(3)) /= 6) then
    print *, "FAIL: want [6] got [", g(3), "]"
    stop 1
end if
    if ((nint(g(1.5))) /= 5) then
    print *, "FAIL: want [5] got [", nint(g(1.5)), "]"
    stop 1
end if
end program test_generic_ambiguity_function_like_dispatch_with_result_types
