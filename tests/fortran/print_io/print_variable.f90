! vybe-test: fortran/print_io/print_variable
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
integer :: x = 99
if ((x) /= 99) then
    print *, "FAIL: want [99] got [", x, "]"
    stop 1
end if
end program t
