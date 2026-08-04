! vybe-test: fortran/io_file_position/fio_open_stream_position_append_specifier
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    open(10, file='fio_open_append.dat', access='stream', form='unformatted', &
         status='replace', position='append')
    write(10) 1
    close(10, status='delete')
end program t
