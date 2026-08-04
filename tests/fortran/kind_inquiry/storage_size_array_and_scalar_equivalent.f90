! vybe-test: fortran/kind_inquiry/storage_size_array_and_scalar_equivalent
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: s = 0
integer :: a(4) = [1,2,3,4]
print *, storage_size(s), storage_size(a)
end program t
