! vybe-test: fortran/transfer_extended/transfer_array_kind8_to_kind4_pair
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=8) :: big = 1000000000000_8
integer(kind=4) :: parts(2)
parts = transfer(big, parts)
if ((transfer(parts, 0_8)) /= 1000000000000) then
    print *, "FAIL: want [1000000000000] got [", transfer(parts, 0_8), "]"
    stop 1
end if
end program t
