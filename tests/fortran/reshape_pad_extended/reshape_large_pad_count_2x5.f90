! vybe-test: fortran/reshape_pad_extended/reshape_large_pad_count_2x5
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(1) = [7]
integer :: m(2,5)
m = reshape(a, [2, 5], pad=[0])
if ((m(1,1)) /= 7) then
    print *, "FAIL: want [7] got [", m(1,1), "]"
    stop 1
end if
if ((count(m == 0)) /= 9) then
    print *, "FAIL: want [9] got [", count(m == 0), "]"
    stop 1
end if
end program t
