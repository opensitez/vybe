! vybe-test: fortran/allocation_status/allocation_status_08
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
logical, allocatable :: a(:)
integer :: st
allocate(a(2), stat=st)
end program p
