! vybe-test: fortran/types/module_with_function_runtime
! origin: languages/fortran/tests/fortran/test_types.rs

module math_utils
    implicit none
contains
    function square(x) result(res)
        real, intent(in) :: x
        real :: res
        res = x * x
    end function square
end module math_utils

program test
    use math_utils
    if ((nint(square(5.0))) /= 25) then
    print *, "FAIL: want [25] got [", nint(square(5.0)), "]"
    stop 1
end if
end program test
