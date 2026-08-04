! vybe-test: fortran/include_directive_line_continuations/test_include_directive_line_continuations_use_ampersand
! origin: languages/fortran/tests/fortran/test_include_directive_line_continuations.rs

program test_include_directive_line_continuations
    integer :: value
    value = 1 + 2 + &
            3 + 4
    if ((value) /= 10) then
    print *, "FAIL: want [10] got [", value, "]"
    stop 1
end if
end program test_include_directive_line_continuations
