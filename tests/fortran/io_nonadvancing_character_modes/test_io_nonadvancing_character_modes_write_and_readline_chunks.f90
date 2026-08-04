! vybe-test: fortran/io_nonadvancing_character_modes/test_io_nonadvancing_character_modes_write_and_readline_chunks
! origin: languages/fortran/tests/fortran/test_io_nonadvancing_character_modes.rs

program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=40) :: txt
    integer :: n
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'A'
    write(unit, '(A)', advance='no') 'B'
    rewind(unit)
    read(unit, '(A)') txt
    n = len_trim(txt)
    print *, n
    close(unit)
end program test_io_nonadvancing_character_modes
