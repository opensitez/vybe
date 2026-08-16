! vybe-test: fortran/file_positioning_and_backspace/test_file_positioning_and_backspace_moves_cursor
! origin: languages/fortran/tests/fortran/test_file_positioning_and_backspace.rs

program test_file_positioning_and_backspace
    integer :: unit
    integer :: v1
    integer :: v2
    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(I0)') 12
    write(unit, '(I0)') 34
    backspace(unit)
    read(unit, *) v1
    rewind(unit)
    read(unit, *) v2
    print *, v1
    print *, v2
    close(unit)
end program test_file_positioning_and_backspace
