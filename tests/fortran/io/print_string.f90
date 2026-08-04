! vybe-test: fortran/io/print_string
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    if (trim("Hello") /= "Hello") then
    print *, "FAIL: want [Hello] got [", "Hello", "]"
    stop 1
end if
end program test
