! vybe-test: fortran/transfer_extended/transfer_char_four_abcd_roundtrip
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=4) :: s = 'ABCD', u
integer :: n
n = transfer(s, 0)
u = transfer(n, '    ')
if ((u == s) /= 1) then
    print *, "FAIL: want [1] got [", u == s, "]"
    stop 1
end if
end program t
