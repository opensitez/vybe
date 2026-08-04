! vybe-test: fortran/bit_btest_ishftc/ibset_then_iand_same_bit_position
! origin: languages/fortran/tests/fortran/test_bit_btest_ishftc.rs
program t
integer :: x
x = ibset(0, 5)
if ((x) /= 32) then
    print *, "FAIL: want [32] got [", x, "]"
    stop 1
end if
if ((iand(x, ishft(1, 5))) /= 32) then
    print *, "FAIL: want [32] got [", iand(x, ishft(1, 5)), "]"
    stop 1
end if
end program t
