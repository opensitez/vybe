! vybe-test: fortran/format_descriptors/desc_2x_before_integer
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, 2X, I0)', 'n', 42
end program t
