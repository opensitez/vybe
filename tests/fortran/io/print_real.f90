! vybe-test: fortran/io/print_real
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    if (abs((3.14) - 3.14) > 1.0e-6) then
    print *, "FAIL: want [3.14] got [", 3.14, "]"
    stop 1
end if
end program test
