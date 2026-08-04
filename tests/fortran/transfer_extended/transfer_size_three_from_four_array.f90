! vybe-test: fortran/transfer_extended/transfer_size_three_from_four_array
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(4) = [10, 20, 30, 40]
integer :: b(3)
b = transfer(a, b, 3)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 20) then
    print *, "FAIL: want [20] got [", b(2), "]"
    stop 1
end if
if ((b(3)) /= 30) then
    print *, "FAIL: want [30] got [", b(3), "]"
    stop 1
end if
end program t
