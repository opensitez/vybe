! vybe-test: fortran/allocation/alloc_ptr_target_18
! origin: languages/fortran/tests/fortran/test_allocation.rs
program p
integer, pointer :: p(:)
allocate(p(3))
end program p
