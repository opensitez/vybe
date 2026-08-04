! vybe-test: fortran/format_descriptors/desc_10x_decuple_blank
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, 10X, A)', 'x', 'y'
end program t
