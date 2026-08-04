! vybe-test: fortran/io_advanced/write_fmt_tab
! origin: languages/fortran/tests/fortran/test_io_advanced.rs
program t
  write(*, '(A, 5X, A)') 'left', 'right'
end program t
