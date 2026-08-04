! vybe-test: fortran/allocation/dealloc_stmt_02
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a)
end program p
