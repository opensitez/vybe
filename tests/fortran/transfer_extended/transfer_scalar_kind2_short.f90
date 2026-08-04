! vybe-test: fortran/transfer_extended/transfer_scalar_kind2_short
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=2) :: s = 32000_2
integer :: n
n = transfer(s, 0)
if ((n) /= 32000) then
    print *, "FAIL: want [32000] got [", n, "]"
    stop 1
end if
end program t
