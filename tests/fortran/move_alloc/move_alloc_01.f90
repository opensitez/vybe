! vybe-test: fortran/move_alloc/move_alloc_01
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
program p
integer, allocatable :: a(:), b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
