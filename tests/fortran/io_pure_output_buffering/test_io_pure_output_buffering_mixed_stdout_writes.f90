! vybe-test: fortran/io_pure_output_buffering/test_io_pure_output_buffering_mixed_stdout_writes
! origin: languages/fortran/tests/fortran/test_io_pure_output_buffering.rs

program test_io_pure_output_buffering
    write(*, '(I0)') 7
    write(*, '(A)', advance='no') 'x'
    write(*, '(A)') 'y'
end program test_io_pure_output_buffering
