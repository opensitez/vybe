! vybe-test: fortran/format_descriptors/desc_2l3_both_logical_values
! origin: languages/fortran/tests/fortran/test_format_descriptors.rs
program t
print '(2L3)', .true., .false.
end program t
