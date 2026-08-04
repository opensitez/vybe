! vybe-test: fortran/io/print_expression
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    integer :: a, b
    a = 10
    b = 20
    if ((a + b) /= 30) then
    print *, "FAIL: want [30] got [", a + b, "]"
    stop 1
end if
end program test
