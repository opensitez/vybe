! vybe-test: fortran/arithmetic/add_integers
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((3 + 4) /= 7) then
    print *, "FAIL: want [7] got [", 3 + 4, "]"
    stop 1
end if
end program t
