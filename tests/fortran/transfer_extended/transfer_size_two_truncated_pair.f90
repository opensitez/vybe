! vybe-test: fortran/transfer_extended/transfer_size_two_truncated_pair
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: b(2)
b = transfer(a, b, 2)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 2) then
    print *, "FAIL: want [2] got [", b(2), "]"
    stop 1
end if
end program t
