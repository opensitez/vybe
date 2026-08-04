! vybe-test: fortran/reshape_pad_extended/reshape_fortran_2x3_negative_values
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [-1, 2, -3, 4, -5, 6]
integer :: m(2,3)
m = reshape(a, [2, 3])
if ((m(1,1)) /= -1) then
    print *, "FAIL: want [-1] got [", m(1,1), "]"
    stop 1
end if
if ((m(2,3)) /= 6) then
    print *, "FAIL: want [6] got [", m(2,3), "]"
    stop 1
end if
end program t
