! vybe-test: fortran/move_alloc/move_alloc_09
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
program p
integer, allocatable :: a(:)
allocate(a(0))
call move_alloc(a,a)
end program p
