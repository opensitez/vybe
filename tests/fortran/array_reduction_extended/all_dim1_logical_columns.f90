! vybe-test: fortran/array_reduction_extended/all_dim1_logical_columns
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(2,3) = reshape([.true.,.true.,.true.,.false.,.true.,.true.],[2,3])
logical :: c(3)
c = all(m, dim=1)
if ((c(1)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", c(1), "]"
    stop 1
end if
if ((c(2)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", c(2), "]"
    stop 1
end if
if ((c(3)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", c(3), "]"
    stop 1
end if
end program t
