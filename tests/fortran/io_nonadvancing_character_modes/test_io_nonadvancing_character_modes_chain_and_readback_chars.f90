! vybe-test: fortran/io_nonadvancing_character_modes/test_io_nonadvancing_character_modes_chain_and_readback_chars
! origin: languages/fortran/tests/fortran/test_io_nonadvancing_character_modes.rs

program test_io_nonadvancing_character_modes
    integer :: unit
    character(len=10) :: a, b
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'x'
    write(unit, '(A)', advance='no') 'y'
    write(unit, '(A)', advance='no') 'z'
    rewind(unit)
    read(unit, '(A1)') a
    read(unit, '(A1)') b
    print *, a
    print *, b
    close(unit)
end program test_io_nonadvancing_character_modes
