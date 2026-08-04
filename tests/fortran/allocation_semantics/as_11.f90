! vybe-test: fortran/allocation_semantics/as_11
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
