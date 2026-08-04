! vybe-test: fortran/fortran2003_extended/runtime_stream_mixed_type_roundtrip
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: u
    integer :: i
    real :: r
    open(newunit=u, status='scratch', access='stream', form='unformatted')
    write(u) 4, 2.5
    rewind(u)
    read(u) i, r
    close(u)
    print *, i
    print *, int(r)
end program t
