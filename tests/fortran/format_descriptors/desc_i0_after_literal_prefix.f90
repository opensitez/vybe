! vybe-test: fortran/format_descriptors/desc_i0_after_literal_prefix
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '("n=", I0)', 99
end program t
