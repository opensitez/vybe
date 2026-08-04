! vybe-test: fortran/allocation/move_alloc_06
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:), b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
