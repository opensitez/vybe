! vybe-test: fortran/allocation_status/allocation_status_02
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3))
deallocate(a, stat=st)
end program p
