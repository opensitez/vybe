! vybe-test: fortran/reshape_pad_extended/reshape_fortran_3x2_sum_check
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(6) = [2, 2, 2, 2, 2, 2]
integer :: m(3,2)
m = reshape(a, [3, 2])
if ((sum(m)) /= 12) then
    print *, "FAIL: want [12] got [", sum(m), "]"
    stop 1
end if
end program t
