! vybe-test: fortran/subroutine_extended/function_real_average_two_values
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((avg2(6.0, 4.0)) /= 5) then
    print *, "FAIL: want [5] got [", avg2(6.0, 4.0), "]"
    stop 1
end if
contains
real function avg2(a, b)
real, intent(in) :: a, b
avg2 = (a + b) / 2.0
end function avg2
end program t
