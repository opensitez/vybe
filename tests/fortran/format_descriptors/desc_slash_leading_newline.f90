! vybe-test: fortran/format_descriptors/desc_slash_leading_newline
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(/, A)', 'alone'
end program t
