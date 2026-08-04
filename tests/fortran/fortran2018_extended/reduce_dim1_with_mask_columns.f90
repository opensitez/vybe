! vybe-test: fortran/fortran2018_extended/reduce_dim1_with_mask_columns
! origin: languages/fortran/tests/fortran/test_fortran2018_extended.rs
program t
integer :: m(2,3) = reshape([1,2,3,4,5,6],[2,3])
logical :: mask(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
integer :: r(3)
r = reduce(m, operator(+), dim=1, mask=mask)
if ((r(1)) /= 1) then
    print *, "FAIL: want [1] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 5) then
    print *, "FAIL: want [5] got [", r(2), "]"
    stop 1
end if
if ((r(3)) /= 3) then
    print *, "FAIL: want [3] got [", r(3), "]"
    stop 1
end if
end program t
