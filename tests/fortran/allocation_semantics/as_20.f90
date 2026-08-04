! vybe-test: fortran/allocation_semantics/as_20
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program p
class(*), allocatable :: x
allocate(integer :: x)
end program p
