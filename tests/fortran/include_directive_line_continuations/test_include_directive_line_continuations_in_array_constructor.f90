! vybe-test: fortran/include_directive_line_continuations/test_include_directive_line_continuations_in_array_constructor
! origin: languages/fortran/tests/fortran/test_include_directive_line_continuations.rs

program test_include_directive_line_continuations
    integer :: a(3)
    a = [ 1,  &
          2,  &
          3 ]
    if ((sum(a)) /= 6) then
    print *, "FAIL: want [6] got [", sum(a), "]"
    stop 1
end if
end program test_include_directive_line_continuations
