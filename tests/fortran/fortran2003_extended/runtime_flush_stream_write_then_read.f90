! vybe-test: fortran/fortran2003_extended/runtime_flush_stream_write_then_read
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: u
    integer :: value

    open(newunit=u, status='scratch')
    write(u, '(I0)') 123
    flush(u)
    rewind(u)
    read(u, '(I0)') value
    close(u)
    print *, value
end program t
