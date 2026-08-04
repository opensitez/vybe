! vybe-test: fortran/fortran2003_extended/compile_flush_scratch_unit_after_write
! origin: languages/fortran/tests/fortran/test_fortran2003_extended.rs

program t
    integer :: u = 10
    open(u, status='scratch')
    write(u, *) 42
    flush(u)
    close(u)
    print *, 'done'
end program t
