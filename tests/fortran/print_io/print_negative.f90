! vybe-test: fortran/print_io/print_negative
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((-10) /= -10) then
    print *, "FAIL: want [-10] got [", -10, "]"
    stop 1
end if
end program t
