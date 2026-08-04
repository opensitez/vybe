! vybe-test: fortran/allocation_status/allocation_status_03
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
integer, allocatable :: x
integer :: st
allocate(x, stat=st)
end program p
