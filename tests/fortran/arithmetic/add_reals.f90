! vybe-test: fortran/arithmetic/add_reals
! origin: languages/fortran/tests/fortran/test_arithmetic.rs
program t
if ((1.5 + 2.5) /= 4) then
    print *, "FAIL: want [4] got [", 1.5 + 2.5, "]"
    stop 1
end if
end program t
