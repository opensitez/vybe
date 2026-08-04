! vybe-test: fortran/transfer_extended/transfer_array_kind1_four_bytes
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer(kind=1) :: b(4) = [18_1, 52_1, 86_1, 120_1]
integer :: n
n = transfer(b, 0)
if ((n) /= 305419896) then
    print *, "FAIL: want [305419896] got [", n, "]"
    stop 1
end if
end program t
