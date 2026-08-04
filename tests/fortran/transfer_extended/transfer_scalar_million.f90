! vybe-test: fortran/transfer_extended/transfer_scalar_million
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: i = 1000000
if ((transfer(i, 0)) /= 1000000) then
    print *, "FAIL: want [1000000] got [", transfer(i, 0), "]"
    stop 1
end if
end program t
