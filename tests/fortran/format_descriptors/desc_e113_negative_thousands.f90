! vybe-test: fortran/format_descriptors/desc_e113_negative_thousands
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(E11.3)', -4.5e3
end program t
