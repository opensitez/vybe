! vybe-test: fortran/transfer/transfer_byte_array_roundtrip
! origin: languages/fortran/tests/fortran/test_transfer.rs

program test
    integer :: original = 305419896
    integer(kind=1) :: bytes(4)
    integer :: recovered
    bytes = transfer(original, bytes)
    recovered = transfer(bytes, 0)
    if (.not. (original == recovered)) then
    print *, "FAIL: want [1] got [", original == recovered, "]"
    stop 1
end if
end program test
