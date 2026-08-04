! vybe-test: fortran/print_io/print_explicit_format_omits_descriptor_literal
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
print "(a, i0)", "Tree size = ", 12
end program t
