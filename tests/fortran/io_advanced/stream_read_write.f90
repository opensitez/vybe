! vybe-test: fortran/io_advanced/stream_read_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: x, y
    open(unit=10, file='stream.bin', access='stream', form='unformatted', &
         status='replace')
    write(10) 100, 200
    rewind(10)
    read(10) x, y
    close(10)
    print *, x, y
end program test
