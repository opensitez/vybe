! vybe-test: fortran/reshape_pad_extended/reshape_repeated_values_pattern
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(8) = [2, 2, 2, 2, 2, 2, 2, 2]
integer :: m(2,2,2)
m = reshape(a, [2, 2, 2])
if ((sum(m)) /= 16) then
    print *, "FAIL: want [16] got [", sum(m), "]"
    stop 1
end if
end program t
