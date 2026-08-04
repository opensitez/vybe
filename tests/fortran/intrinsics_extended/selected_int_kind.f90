! vybe-test: fortran/intrinsics_extended/selected_int_kind
! origin: languages/fortran/tests/fortran/test_intrinsics_extended.rs
program t
integer :: k
k = selected_int_kind(9)
print *, k
end program t
