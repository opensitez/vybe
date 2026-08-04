! vybe-test: fortran/print_io/max_inside_addition_keeps_outer_operand
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((1 + max(0, 0)) /= 1) then
    print *, "FAIL: want [1] got [", 1 + max(0, 0), "]"
    stop 1
end if
end program t
