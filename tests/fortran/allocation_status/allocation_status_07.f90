! vybe-test: fortran/allocation_status/allocation_status_07
! origin: languages/fortran/tests/fortran/test_allocation_status.rs
program p
class(*), allocatable :: x
integer :: st
allocate(integer :: x, stat=st)
end program p
