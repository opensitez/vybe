! vybe-test: fortran/transfer/transfer_char_to_int
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    character(len=4) :: s
    integer :: n
    s = transfer(0, s)
    n = transfer(s, 0)
    if ((n) /= 0) then
    print *, "FAIL: want [0] got [", n, "]"
    stop 1
end if
end program test
