! vybe-test: fortran/transfer_extended/transfer_array_logical_to_integer
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
logical :: m(2) = [.true., .false.]
integer :: n(2)
n = transfer(m, n)
if ((n(1)) /= 1) then
    print *, "FAIL: want [1] got [", n(1), "]"
    stop 1
end if
if ((n(2)) /= 0) then
    print *, "FAIL: want [0] got [", n(2), "]"
    stop 1
end if
end program t
