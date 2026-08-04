! vybe-test: fortran/bit_btest_ishftc/not_negative_two
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((not(-2)) /= 1) then
    print *, "FAIL: want [1] got [", not(-2), "]"
    stop 1
end if
end program t
