! vybe-test: fortran/bit_btest_ishftc/iand_negative_one_with_byte_mask
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
if ((iand(-1, 255)) /= 255) then
    print *, "FAIL: want [255] got [", iand(-1, 255), "]"
    stop 1
end if
end program t
