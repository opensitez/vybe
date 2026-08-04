! vybe-test: fortran/format_io_extended/fmt_label_write_logical
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
logical :: flag = .false.
write(*, 400) flag
400 format(L5)
end program t
