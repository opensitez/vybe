! vybe-test: fortran/reshape_pad_extended/reshape_fortran_1x6_row_vector
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [10, 20, 30, 40, 50, 60]
integer :: m(1,6)
m = reshape(a, [1, 6])
if ((m(1,1)) /= 10) then
    print *, "FAIL: want [10] got [", m(1,1), "]"
    stop 1
end if
if ((m(1,6)) /= 60) then
    print *, "FAIL: want [60] got [", m(1,6), "]"
    stop 1
end if
end program t
