! vybe-test: fortran/transfer_extended/transfer_char_two_bytes_ab
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=2) :: s = 'AB'
if ((transfer(s, 0)) /= 16961) then
    print *, "FAIL: want [16961] got [", transfer(s, 0), "]"
    stop 1
end if
end program t
