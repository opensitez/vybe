! vybe-test: fortran/io_advanced/io_position_and_backspace_reseek
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: a, b
    open(unit=10, file='adv_pos.bin', status='replace', action='readwrite')
    write(10, '(I0)') 11
    write(10, '(I0)') 22
    backspace(10)
    read(10, *) a
    rewind(10)
    read(10, *) b
    close(10)
    print *, a + b
end program test
