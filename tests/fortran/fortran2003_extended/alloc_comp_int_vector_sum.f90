! vybe-test: fortran/fortran2003_extended/alloc_comp_int_vector_sum
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Vec
integer, allocatable :: data(:)
end type Vec
type(Vec) :: v
v%data = [2, 4, 6]
if ((sum(v%data)) /= 12) then
    print *, "FAIL: want [12] got [", sum(v%data), "]"
    stop 1
end if
end program t
