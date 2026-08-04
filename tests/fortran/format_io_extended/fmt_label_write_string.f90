! vybe-test: fortran/format_io_extended/fmt_label_write_string
! origin: languages/fortran/tests/fortran/test_format_io_extended.rs
program t
character(len=5) :: s = 'hello'
write(*, 300) s
300 format(A)
end program t
