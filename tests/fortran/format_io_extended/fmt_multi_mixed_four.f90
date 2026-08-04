! vybe-test: fortran/format_io_extended/fmt_multi_mixed_four
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(A, I0, F5.1, L5)', 'v', 3, 1.4, .false.
end program t
