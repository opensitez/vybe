! vybe-test: fortran/else_if_cascade_priority/test_else_if_cascade_short_circuits_after_first_match
! origin: languages/fortran/tests/fortran/test_else_if_cascade_priority.rs

program test_else_if_cascade_priority_short_circuit
    integer :: x, y
    x = 9
    y = 0
    if (x > 5) then
        y = 10
    else if (x > 7) then
        y = 20
    else if (x > 8) then
        y = 30
    else
        y = 40
    end if
    if ((y) /= 10) then
    print *, "FAIL: want [10] got [", y, "]"
    stop 1
end if
end program test_else_if_cascade_priority_short_circuit
