! vybe-test: fortran/reshape_pad_extended/reshape_from_array_constructor
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: m(2,2)
m = reshape([(i, i = 1, 4)], [2, 2])
if ((m(2,2)) /= 4) then
    print *, "FAIL: want [4] got [", m(2,2), "]"
    stop 1
end if
end program t
