! vybe-test: fortran/file_rewind_and_endfile_behavior/test_file_rewind_and_endfile_behavior
! origin: languages/fortran/tests/fortran/test_file_rewind_and_endfile_behavior.rs

program test_file_rewind_and_endfile_behavior
    integer :: unit
    integer :: first
    integer :: second
    integer :: code

    open(newunit=unit, status='scratch', action='readwrite')
    write(unit, '(I0)') 7
    write(unit, '(I0)') 9
    rewind(unit)
    read(unit, *) first
    endfile(unit)
    read(unit, *, iostat=code) second

    print *, first
    print *, code
    close(unit)
end program test_file_rewind_and_endfile_behavior
