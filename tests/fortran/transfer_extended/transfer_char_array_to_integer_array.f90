! vybe-test: fortran/transfer_extended/transfer_char_array_to_integer_array
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
character(len=1) :: c(3) = ['A', 'B', 'C']
integer :: n(3)
n = transfer(c, n)
if ((n(1)) /= 65) then
    print *, "FAIL: want [65] got [", n(1), "]"
    stop 1
end if
if ((n(2)) /= 66) then
    print *, "FAIL: want [66] got [", n(2), "]"
    stop 1
end if
if ((n(3)) /= 67) then
    print *, "FAIL: want [67] got [", n(3), "]"
    stop 1
end if
end program t
