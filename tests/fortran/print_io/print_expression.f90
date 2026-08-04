! vybe-test: fortran/print_io/print_expression
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((2 + 3) /= 5) then
    print *, "FAIL: want [5] got [", 2 + 3, "]"
    stop 1
end if
end program t
