! vybe-test: fortran/allocation/alloc_nullify_ptr_30
! origin: languages/fortran/tests/fortran/test_allocation.rs
program driver
integer, pointer :: p(:)
nullify(p)
end program driver