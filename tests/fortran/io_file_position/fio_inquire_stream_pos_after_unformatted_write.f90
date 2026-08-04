! vybe-test: fortran/io_file_position/fio_inquire_stream_pos_after_unformatted_write
! origin: languages/fortran/tests/fortran/test_io_file_position.rs

program t
    integer :: pos
    open(10, status='scratch', access='stream', form='unformatted')
    write(10) 1, 2, 3
    inquire(unit=10, pos=pos)
    close(10)
    print *, pos
end program t
