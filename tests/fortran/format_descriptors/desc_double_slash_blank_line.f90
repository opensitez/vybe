! vybe-test: fortran/format_descriptors/desc_double_slash_blank_line
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, /, /, A)', 'a', 'b'
end program t
