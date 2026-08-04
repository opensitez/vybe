! vybe-test: fortran/format_io_extended/fmt_multi_a_i_l
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(A, I0, L5)', 'flag=', 1, .true.
end program t
