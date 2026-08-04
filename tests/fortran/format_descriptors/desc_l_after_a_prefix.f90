! vybe-test: fortran/format_descriptors/desc_l_after_a_prefix
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(A, L5)', 'ok=', .true.
end program t
