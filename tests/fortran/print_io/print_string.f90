! vybe-test: fortran/print_io/print_string
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if (trim("Hello") /= "Hello") then
    print *, "FAIL: want [Hello] got [", "Hello", "]"
    stop 1
end if
end program t
