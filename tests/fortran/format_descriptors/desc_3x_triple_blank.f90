! vybe-test: fortran/format_descriptors/desc_3x_triple_blank
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, 3X, A)', 'L', 'R'
end program t
