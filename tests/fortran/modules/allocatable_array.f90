! vybe-test: fortran/modules/allocatable_array
! origin: languages/fortran/tests/fortran/test_modules.rs
program t
integer, allocatable :: arr(:)
allocate(arr(5))
arr(1) = 42
print *, arr(1)
deallocate(arr)
end program t
