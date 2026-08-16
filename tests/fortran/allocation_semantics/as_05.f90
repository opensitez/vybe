! vybe-test: fortran/allocation_semantics/as_05
! origin: languages/fortran/tests/fortran/test_allocation_semantics.rs
program driver
integer, pointer :: p(:)
allocate(p(3))
end program driver