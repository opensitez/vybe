! vybe-test: fortran/io_file_position/fio_open_stream_position_rewind_specifier
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    integer :: v
    open(10, file='fio_open_rewind.dat', access='stream', form='unformatted', &
         status='replace', position='rewind')
    write(10) 12
    rewind(10)
    read(10) v
    close(10, status='delete')
    print *, v
end program t
