! vybe-test: fortran/io_nonadvancing_character_modes/test_io_nonadvancing_character_modes_read_iostat
! origin: languages/fortran/tests/fortran/test_io_nonadvancing_character_modes.rs

program test_io_nonadvancing_character_modes
    integer :: ios
    integer :: unit
    character(len=2) :: token
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(A)', advance='no') 'q'
    rewind(unit)
    read(unit, '(A2)', iostat=ios) token
    print *, trim(token)
    if (ios /= 0) then
        print *, 1
    else
        print *, 0
    end if
    close(unit)
end program test_io_nonadvancing_character_modes
