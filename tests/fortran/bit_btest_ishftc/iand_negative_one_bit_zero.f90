! vybe-test: fortran/bit_btest_ishftc/iand_negative_one_bit_zero
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((iand(-1, ishft(1, 0))) /= 1) then
    print *, "FAIL: want [1] got [", iand(-1, ishft(1, 0)), "]"
    stop 1
end if
end program t
