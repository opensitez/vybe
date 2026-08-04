! vybe-test: fortran/array_reduction_extended/any_dim2_logical_rows
! origin: languages/fortran/tests/fortran/test_array_reduction_extended.rs
program t
logical :: m(3,2) = reshape([.false.,.true.,.false.,.false.,.false.,.false.],[3,2])
logical :: r(3)
r = any(m, dim=2)
if ((r(1)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", r(1), "]"
    stop 1
end if
if ((r(2)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", r(2), "]"
    stop 1
end if
if ((r(3)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", r(3), "]"
    stop 1
end if
end program t
