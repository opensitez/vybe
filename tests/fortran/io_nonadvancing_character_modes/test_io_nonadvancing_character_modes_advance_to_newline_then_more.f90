! vybe-test: fortran/io_nonadvancing_character_modes/test_io_nonadvancing_character_modes_advance_to_newline_then_more
! origin: languages/fortran/tests/fortran/test_io_nonadvancing_character_modes.rs

program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=20) :: line
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'left'
    write(unit, '(A)', advance='yes') ','
    write(unit, '(A)', advance='no') 'right'
    rewind(unit)
    read(unit, '(A)') line
    print *, trim(line)
    close(unit)
end program test_io_nonadvancing_character_modes
