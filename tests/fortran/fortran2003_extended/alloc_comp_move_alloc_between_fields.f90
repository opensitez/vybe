! vybe-test: fortran/fortran2003_extended/alloc_comp_move_alloc_between_fields
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs
program t
type :: Pair
integer, allocatable :: left(:), right(:)
end type Pair
type(Pair) :: p
allocate(p%left(2))
p%left = [5, 6]
call move_alloc(p%left, p%right)
if ((p%right(1)) /= 5) then
    print *, "FAIL: want [5] got [", p%right(1), "]"
    stop 1
end if
if ((allocated(p%left))) then
    print *, "FAIL: want [0] got [", allocated(p%left), "]"
    stop 1
end if
end program t
