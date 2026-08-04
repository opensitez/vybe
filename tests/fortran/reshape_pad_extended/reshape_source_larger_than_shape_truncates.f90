! vybe-test: fortran/reshape_pad_extended/reshape_source_larger_than_shape_truncates
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(10) = [(i, i = 1, 10)]
integer :: m(2,2)
m = reshape(a, [2, 2])
if ((sum(m)) /= 10) then
    print *, "FAIL: want [10] got [", sum(m), "]"
    stop 1
end if
end program t
