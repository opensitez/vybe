! vybe-test: fortran/stream_access_io/compile_stream_position_inquire
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: pos
    open(31, status='scratch', access='stream', form='unformatted')
    write(31) 1, 2, 3
    inquire(31, pos=pos)
    close(31)
    print *, pos
end program t
