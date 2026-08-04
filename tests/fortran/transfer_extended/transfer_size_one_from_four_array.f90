! vybe-test: fortran/transfer_extended/transfer_size_one_from_four_array
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(4) = [10, 20, 30, 40]
integer :: b(1)
b = transfer(a, b, 1)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
end program t
