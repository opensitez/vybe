! vybe-test: fortran/print_io/print_integer
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((42) /= 42) then
    print *, "FAIL: want [42] got [", 42, "]"
    stop 1
end if
end program t
