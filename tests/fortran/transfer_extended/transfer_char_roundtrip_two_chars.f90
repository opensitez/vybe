! vybe-test: fortran/transfer_extended/transfer_char_roundtrip_two_chars
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=2) :: s = 'Hi', t
integer :: n
n = transfer(s, 0)
t = transfer(n, '  ')
if ((ichar(t(1:1))) /= 72) then
    print *, "FAIL: want [72] got [", ichar(t(1:1)), "]"
    stop 1
end if
if ((ichar(t(2:2))) /= 105) then
    print *, "FAIL: want [105] got [", ichar(t(2:2)), "]"
    stop 1
end if
end program t
