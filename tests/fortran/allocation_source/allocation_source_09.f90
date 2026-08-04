! vybe-test: fortran/allocation_source/allocation_source_09
! origin: languages/fortran/tests/fortran/test_allocation_source.rs
program p
class(*), allocatable :: x
allocate(integer :: x, source=1)
end program p
