! vybe-test: fortran/io_pure_output_buffering/test_io_pure_output_buffering_buffered_file_roundtrip
! origin: languages/fortran/tests/fortran/test_io_pure_output_buffering.rs

program test_io_pure_output_buffering
    integer :: unit
    character(len=10) :: a
    open(newunit=unit, file='pure_buf.txt', status='replace', action='readwrite')
    write(unit, '(I0)') 123
    write(unit, '(A)', advance='no') 'ab'
    write(unit, '(A)') 'cd'
    rewind(unit)
    read(unit, '(A)') a
    close(unit, status='delete')
    print *, trim(a)
end program test_io_pure_output_buffering
