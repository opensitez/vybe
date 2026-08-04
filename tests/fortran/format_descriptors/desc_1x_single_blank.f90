! vybe-test: fortran/format_descriptors/desc_1x_single_blank
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, 1X, A)', 'a', 'b'
end program t
