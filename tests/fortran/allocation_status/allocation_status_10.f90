! vybe-test: fortran/allocation_status/allocation_status_10
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
real, allocatable :: a(:,:)
integer :: st
allocate(a(2,2), stat=st)
end program p
