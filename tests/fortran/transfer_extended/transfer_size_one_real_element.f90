! vybe-test: fortran/transfer_extended/transfer_size_one_real_element
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
real :: a(3) = [1.0, 2.0, 3.0]
real :: b(1)
b = transfer(a, b, 1)
if ((b(1)) /= 1) then
    print *, "FAIL: want [1] got [", b(1), "]"
    stop 1
end if
end program t
