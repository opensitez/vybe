! vybe-test: fortran/io/print_integer
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    if ((42) /= 42) then
    print *, "FAIL: want [42] got [", 42, "]"
    stop 1
end if
end program test
