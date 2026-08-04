! vybe-test: fortran/transfer_extended/transfer_array_int_four_elements
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: b(4)
b = transfer(a, b)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
if ((b(4)) /= 4) then
    print *, "FAIL: want [4] got [", b(4), "]"
    stop 1
end if
end program t
