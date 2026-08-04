! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_keyword_defaults
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program subroutine_argument_default_values_keyword_defaults
    if ((combine(a=1)) /= 1) then
    print *, "FAIL: want [1] got [", combine(a=1), "]"
    stop 1
end if
    if ((combine(a=1, b=2)) /= 3) then
    print *, "FAIL: want [3] got [", combine(a=1, b=2), "]"
    stop 1
end if
contains
    integer function combine(a, b, c)
        integer, intent(in) :: a
        integer, intent(in), optional :: b, c
        combine = a
        if (present(b)) combine = combine + b
        if (present(c)) combine = combine + c
    end function combine
end program subroutine_argument_default_values_keyword_defaults
