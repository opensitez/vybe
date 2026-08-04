! vybe-test: fortran/allocation/deferred_shape_09
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, allocatable :: a(:,:)
allocate(a(2,2))
end program p
