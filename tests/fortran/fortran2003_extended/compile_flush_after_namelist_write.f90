! vybe-test: fortran/fortran2003_extended/compile_flush_after_namelist_write
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: n = 3
    real :: x = 1.5
    open(20, status='scratch')
    write(20, nml=cfg)
    flush(20)
    close(20)
    print *, n
contains
    namelist /cfg/ n, x
end program t
