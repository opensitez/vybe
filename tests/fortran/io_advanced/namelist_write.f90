! vybe-test: fortran/io_advanced/namelist_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: x = 10, y = 20
    real :: z = 3.25
    namelist /cfg/ x, y, z
    open(unit=10, file='namelist_roundtrip.nml', status='replace', action='readwrite')
    write(10, nml=cfg)
    rewind(10)
    x = 0
    y = 0
    z = 0.0
    read(10, nml=cfg)
    close(10)
    print *, x
    print *, y
    print *, int(z * 100.0)
end program test
