! vybe-test: fortran/fortran2003_extended/alloc_comp_real_matrix_corner
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Grid
real, allocatable :: cells(:,:)
end type Grid
type(Grid) :: g
allocate(g%cells(2, 2))
g%cells = reshape([1.0, 2.0, 3.0, 4.0], [2, 2])
if ((int(g%cells(2, 1))) /= 3) then
    print *, "FAIL: want [3] got [", int(g%cells(2, 1)), "]"
    stop 1
end if
end program t
