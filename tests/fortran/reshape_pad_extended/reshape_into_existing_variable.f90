! vybe-test: fortran/reshape_pad_extended/reshape_into_existing_variable
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [1, 2, 3, 4, 5, 6]
integer :: m(2,3)
m = reshape(a, shape=[2, 3])
if ((m(1,2)) /= 3) then
    print *, "FAIL: want [3] got [", m(1,2), "]"
    stop 1
end if
end program t
