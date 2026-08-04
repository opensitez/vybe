! vybe-test: fortran/format_descriptors/desc_f84_negative_value
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(F8.4)', -0.375
end program t
