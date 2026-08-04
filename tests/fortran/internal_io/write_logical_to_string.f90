! vybe-test: fortran/internal_io/write_logical_to_string
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=5) :: buf
    write(buf, '(L5)') .true.
    print *, trim(buf)
end program test
