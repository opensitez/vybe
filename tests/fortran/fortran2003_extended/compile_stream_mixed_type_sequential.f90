! vybe-test: fortran/fortran2003_extended/compile_stream_mixed_type_sequential
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: i
    real :: r
    character(len=4) :: tag
    open(32, status='scratch', access='stream', form='unformatted')
    write(32) 5, 2.5
    rewind(32)
    read(32) i, r
    close(32)
    print *, i, int(r)
end program t
