! vybe-test: fortran/subroutine_argument_default_values/subroutine_argument_default_values_keyword_skip_with_optional_tail
! origin: languages/fortran/tests/fortran/test_subroutine_argument_default_values.rs

program t
    if ((combine(a=2, c=7)) /= 9) then
    print *, "FAIL: want [9] got [", combine(a=2, c=7), "]"
    stop 1
end if
    if ((combine(a=2, b=4, c=7)) /= 13) then
    print *, "FAIL: want [13] got [", combine(a=2, b=4, c=7), "]"
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
end program t
