! vybe-test: fortran/reshape_pad_extended/reshape_pad_larger_than_source_2x4
! origin: languages/fortran/tests/fortran/test_reshape_pad_extended.rs
program t
integer :: a(5) = [1, 2, 3, 4, 5]
integer :: m(2,4)
m = reshape(a, [2, 4], pad=[0])
if ((m(2,4)) /= 0) then
    print *, "FAIL: want [0] got [", m(2,4), "]"
    stop 1
end if
if ((sum(m)) /= 15) then
    print *, "FAIL: want [15] got [", sum(m), "]"
    stop 1
end if
end program t
