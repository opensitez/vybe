! vybe-test: fortran/allocation/alloc_stmt_01
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:)
allocate(a(3))
end program p
