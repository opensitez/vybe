! vybe-test: fortran/move_alloc/move_alloc_07
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
type t
 integer :: x
end type t
program p
type(t), allocatable :: a,b
allocate(a)
call move_alloc(a,b)
end program p
