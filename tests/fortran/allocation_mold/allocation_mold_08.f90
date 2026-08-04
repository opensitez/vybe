! vybe-test: fortran/allocation_mold/allocation_mold_08
! origin: languages/fortran/tests/fortran/test_allocation_mold.rs
program p
class(*), allocatable :: x, y
allocate(integer :: y)
allocate(x, mold=y)
end program p
