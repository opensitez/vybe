! vybe-test: fortran/transfer_extended/transfer_array_2d_flatten_to_1d
! origin: languages/fortran/tests/fortran/test_transfer_extended.rs
program t
integer :: m(2,2) = reshape([1, 2, 3, 4], [2, 2])
integer :: v(4)
v = transfer(m, v)
if ((v(1)) /= 1) then
    print *, "FAIL: want [1] got [", v(1), "]"
    stop 1
end if
if ((v(4)) /= 4) then
    print *, "FAIL: want [4] got [", v(4), "]"
    stop 1
end if
end program t
