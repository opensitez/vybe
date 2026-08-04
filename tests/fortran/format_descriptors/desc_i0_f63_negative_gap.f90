! vybe-test: fortran/format_descriptors/desc_i0_f63_negative_gap
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(I0, F6.3)', 1, -2.5
end program t
