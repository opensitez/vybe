! vybe-test: fortran/allocation_status/allocation_status_06
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program driver
integer, pointer :: p(:)
integer :: st
allocate(p(3), stat=st)
end program driver