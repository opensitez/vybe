! vybe-test: fortran/reshape_pad_extended/reshape_expand_1d_to_2d_count
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 1, 1, 1]
integer :: m(2,2)
m = reshape(a, [2, 2])
if ((count(m == 1)) /= 4) then
    print *, "FAIL: want [4] got [", count(m == 1), "]"
    stop 1
end if
end program t
