! vybe-test: fortran/transfer_extended/transfer_scalar_fortytwo
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 42
if ((transfer(i, 0)) /= 42) then
    print *, "FAIL: want [42] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
