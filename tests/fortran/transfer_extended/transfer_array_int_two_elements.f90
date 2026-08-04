! vybe-test: fortran/transfer_extended/transfer_array_int_two_elements
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: a(2) = [10, 20]
integer :: b(2)
b = transfer(a, b)
if ((b(1)) /= 10) then
    print *, "FAIL: want [10] got [", b(1), "]"
    stop 1
end if
if ((b(2)) /= 20) then
    print *, "FAIL: want [20] got [", b(2), "]"
    stop 1
end if
end program t
