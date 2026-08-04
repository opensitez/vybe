! vybe-test: fortran/reshape_pad_extended/reshape_all_zeros_pad_stays_zero
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(2) = [0, 0]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=5)
if ((count(m == 0)) /= 2) then
    print *, "FAIL: want [2] got [", count(m == 0), "]"
    stop 1
end if
if ((count(m == 5)) /= 2) then
    print *, "FAIL: want [2] got [", count(m == 5), "]"
    stop 1
end if
end program t
