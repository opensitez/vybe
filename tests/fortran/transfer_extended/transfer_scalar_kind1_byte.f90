! vybe-test: fortran/transfer_extended/transfer_scalar_kind1_byte
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=1) :: b = 127_1
integer :: n
n = transfer(b, 0)
if ((n) /= 127) then
    print *, "FAIL: want [127] got [", n, "]"
    stop 1
end if
end program t
