! vybe-test: fortran/reshape_pad_extended/reshape_order_c_2x2_anti_diagonal
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: m(2,2)
m = reshape(a, [2, 2], order='C')
if ((m(1,1)) /= 1) then
    print *, "FAIL: want [1] got [", m(1,1), "]"
    stop 1
end if
if ((m(1,2)) /= 2) then
    print *, "FAIL: want [2] got [", m(1,2), "]"
    stop 1
end if
if ((m(2,1)) /= 3) then
    print *, "FAIL: want [3] got [", m(2,1), "]"
    stop 1
end if
if ((m(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", m(2,2), "]"
    stop 1
end if
end program t
