! vybe-test: fortran/intrinsics_extended/selected_real_kind
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
integer :: k
k = selected_real_kind(15)
print *, k
end program t
