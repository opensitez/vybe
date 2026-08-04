! vybe-test: fortran/allocation_semantics/as_04
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:),b(:)
allocate(a(3))
call move_alloc(a,b)
end program p
