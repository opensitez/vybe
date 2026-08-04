! vybe-test: fortran/io_advise_and_error_state/test_io_advise_and_error_state_endfile_sets_end
! origin: languages/fortran/tests/fortran/test_io_advise_and_error_state.rs

program test_io_advise_and_error_state
    integer :: unit
    integer :: n
    integer :: ios
    character(len=12) :: buf
    open(newunit=unit, file='end_probe.dat', status='replace')
    write(unit, '(I0)') 7
    rewind(unit)
    read(unit, *, iostat=ios) n
    read(unit, *, iostat=ios) n
    if (ios < 0) then
        print *, 1
    else
        print *, 0
    end if
    close(unit, status='delete')
end program test_io_advise_and_error_state
