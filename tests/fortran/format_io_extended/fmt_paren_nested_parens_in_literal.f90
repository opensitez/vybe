! vybe-test: fortran/format_io_extended/fmt_paren_nested_parens_in_literal
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
print '(A, I0)', '(', 99
end program t
