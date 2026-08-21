! vybe-test: fortran/stream_access_io/compile_stream_access_replace_file
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: v
    open(30, file='tmp_stream.bin', access='stream', form='unformatted', status='replace')
    write(30) 77
    rewind(30)
    read(30) v
    close(30, status='delete')
    print *, v
end program t
