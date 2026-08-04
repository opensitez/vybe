! vybe-test: fortran/io_advanced/namelist_read
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: nx = 10, ny = 20
    integer :: ios
    namelist /grid/ nx, ny
    open(unit=11, file='namelist_input.nml', status='replace', action='readwrite')
    write(11, '(A)') '&grid'
    write(11, '(A)') ' nx = 64,'
    write(11, '(A)') ' ny = 32'
    write(11, '(A)') '/'
    rewind(11)
    read(11, nml=grid, iostat=ios)
    close(11)
    print *, ios
    print *, nx
    print *, ny
end program test
