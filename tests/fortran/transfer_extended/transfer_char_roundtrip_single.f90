! vybe-test: fortran/transfer_extended/transfer_char_roundtrip_single
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c = 'X', d
integer :: n
n = transfer(c, 0)
d = transfer(n, ' ')
if ((ichar(d)) /= 88) then
    print *, "FAIL: want [88] got [", ichar(d), "]"
    stop 1
end if
end program t
