! vybe-test: fortran/transfer/transfer_kind8_bytes
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer(kind=8) :: n = 0_8
    integer(kind=1) :: bytes(8)
    bytes = transfer(n, bytes)
    if ((bytes(1)) /= 0) then
    print *, "FAIL: want [0] got [", bytes(1), "]"
    stop 1
end if
end program test
