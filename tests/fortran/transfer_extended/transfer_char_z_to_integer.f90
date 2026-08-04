! vybe-test: fortran/transfer_extended/transfer_char_z_to_integer
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c = 'Z'
if ((transfer(c, 0)) /= 90) then
    print *, "FAIL: want [90] got [", transfer(c, 0), "]"
    stop 1
end if
end program t
