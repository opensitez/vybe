! vybe-test: fortran/io_advanced/stream_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    open(unit=10, file='stream.bin', access='stream', form='unformatted', &
         status='replace')
    write(10) 42
    close(10)
end program test
