! vybe-test: fortran/kind_inquiry/kind_array_vs_scalar_equivalent
! origin: languages/fortran/tests/fortran/test_kind_inquiry.rs
program t
integer :: s = 0
integer :: a(4) = [1,2,3,4]
print *, kind(s), kind(a)
end program t
