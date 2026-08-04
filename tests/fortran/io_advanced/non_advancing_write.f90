! vybe-test: fortran/io_advanced/non_advancing_write
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    write(*, '(A)', advance='no') 'no newline'
    write(*, '(A)') ' here'
end program test
