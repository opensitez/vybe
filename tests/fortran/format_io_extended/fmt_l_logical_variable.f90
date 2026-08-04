! vybe-test: fortran/format_io_extended/fmt_l_logical_variable
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
logical :: ok = .true.
print '(L5)', ok
end program t
