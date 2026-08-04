! vybe-test: fortran/print_io/write_advance_no_buffers_single_line
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
write(*, "(a)", advance="no") "["
write(*, "(i0, a)", advance="no") 3, ", "
write(*, "(i0, a)", advance="no") 5, "]"
write(*, *)
end program t
