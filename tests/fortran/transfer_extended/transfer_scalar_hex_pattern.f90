! vybe-test: fortran/transfer_extended/transfer_scalar_hex_pattern
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 305419896
if ((transfer(i, 0)) /= 305419896) then
    print *, "FAIL: want [305419896] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
