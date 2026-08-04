! vybe-test: fortran/transfer_extended/transfer_size_partial_char_to_int
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=3) :: s = 'abc'
integer :: n
n = transfer(s, 0, 1)
if ((n) /= 97) then
    print *, "FAIL: want [97] got [", n, "]"
    stop 1
end if
end program t
