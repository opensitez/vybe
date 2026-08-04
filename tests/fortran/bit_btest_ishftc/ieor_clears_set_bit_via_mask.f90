! vybe-test: fortran/bit_btest_ishftc/ieor_clears_set_bit_via_mask
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((ieor(255, ishft(1, 4))) /= 239) then
    print *, "FAIL: want [239] got [", ieor(255, ishft(1, 4)), "]"
    stop 1
end if
end program t
