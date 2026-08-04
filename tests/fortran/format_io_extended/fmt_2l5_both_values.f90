! vybe-test: fortran/format_io_extended/fmt_2l5_both_values
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(2L5)', .true., .false.
end program t
