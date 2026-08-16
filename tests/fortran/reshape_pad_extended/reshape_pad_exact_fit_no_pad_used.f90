! vybe-test: fortran/reshape_pad_extended/reshape_pad_exact_fit_no_pad_used
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(4) = [1, 2, 3, 4]
integer :: m(2,2)
m = reshape(a, [2, 2], pad=[99])
if ((sum(m)) /= 10) then
    print *, "FAIL: want [10] got [", sum(m), "]"
    stop 1
end if
end program t
