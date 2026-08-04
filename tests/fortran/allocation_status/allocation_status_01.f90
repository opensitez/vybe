! vybe-test: fortran/allocation_status/allocation_status_01
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
