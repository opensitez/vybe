! vybe-test: fortran/allocation_status/allocation_status_09
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
complex, allocatable :: a(:)
integer :: st
allocate(a(2), stat=st)
end program p
