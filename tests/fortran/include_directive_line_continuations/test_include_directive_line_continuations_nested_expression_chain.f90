! vybe-test: fortran/include_directive_line_continuations/test_include_directive_line_continuations_nested_expression_chain
! origin: languages/fortran/tests/fortran/test_include_directive_line_continuations.rs

program test_include_directive_line_continuations
    integer :: a, b, c
    a = 1 + &
        2 + 3 + &
        4
    b = (a * 2) / &
        (1 + 3)
    c = a + b - 2
    if ((c) /= 13) then
    print *, "FAIL: want [13] got [", c, "]"
    stop 1
end if
end program test_include_directive_line_continuations
