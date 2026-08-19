! vybe-test: fortran/transfer_extended/transfer_size_partial_char_to_int
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=3) :: s = 'abc'
integer :: n(1)
n = transfer(s, 0, 1)
if ((n(1)) /= 6513249) then
    print *, "FAIL: want [6513249] got [", n(1), "]"
    stop 1
end if
end program t
