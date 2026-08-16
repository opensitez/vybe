! vybe-test: fortran/array_reduction_extended/count_dim2_logical_rows
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(2,3) = reshape([.true.,.false.,.true.,.false.,.true.,.false.],[2,3])
integer :: r(2)
r = count(m, dim=2)
if ((r(1)) /= 3) then
    print *, "FAIL: want [3] got [", r(1), "]"
    stop 1
end if
if ((r(2)) /= 0) then
    print *, "FAIL: want [0] got [", r(2), "]"
    stop 1
end if
end program t
