! vybe-test: fortran/transfer/transfer_int_to_char
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: n = 1195853639
    character(len=4) :: s
    s = transfer(n, '    ')
    if ((len(s)) /= 4) then
    print *, "FAIL: want [4] got [", len(s), "]"
    stop 1
end if
end program test
