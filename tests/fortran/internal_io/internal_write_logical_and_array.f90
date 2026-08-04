! vybe-test: fortran/internal_io/internal_write_logical_and_array
! origin: languages/fortran/tests/fortran/test_internal_io.rs

program test
    character(len=20) :: buf
    logical :: flag
    integer :: i
    flag = .true.
    write(buf, '(L1, 1X, I0)') flag, 3
    read(buf, '(L1, 1X, I0)') flag, i
    print *, i
end program test
