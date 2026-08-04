! vybe-test: fortran/transfer_extended/transfer_size_larger_than_source_scalar
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a = 7
integer :: b(8)
b = transfer(a, b, 8)
if ((b(1)) /= 7) then
    print *, "FAIL: want [7] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 0) then
    print *, "FAIL: want [0] got [", b(2), "]"
    stop 1
end if
end program t
