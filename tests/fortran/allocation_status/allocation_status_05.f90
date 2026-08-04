! vybe-test: fortran/allocation_status/allocation_status_05
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
character(len=20) :: msg
integer :: st
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a, stat=st, errmsg=msg)
end program p
