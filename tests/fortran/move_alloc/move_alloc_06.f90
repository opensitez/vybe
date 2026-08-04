! vybe-test: fortran/move_alloc/move_alloc_06
! origin: languages/fortran/tests/fortran/test_move_alloc.rs
program p
complex, allocatable :: a(:), b(:)
allocate(a(2))
call move_alloc(a,b)
end program p
