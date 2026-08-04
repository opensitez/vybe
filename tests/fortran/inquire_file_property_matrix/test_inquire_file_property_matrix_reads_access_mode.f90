! vybe-test: fortran/inquire_file_property_matrix/test_inquire_file_property_matrix_reads_access_mode
! origin: languages/fortran/tests/fortran/test_inquire_file_property_matrix.rs

program test_inquire_file_property_matrix
    logical :: exists
    character(len=16) :: access_mode
    character(len=256) :: path
    path = 'tmp_inquire_access_mode.txt'
    open(unit=99, file=trim(path), status='replace')
    inquire(file=trim(path), exist=exists, access=access_mode)
    if (trim(exists) /= ".TRUE.") then
    print *, "FAIL: want [.TRUE.] got [", exists, "]"
    stop 1
end if
    if (trim(trim(access_mode)) /= "SEQUENTIAL") then
    print *, "FAIL: want [SEQUENTIAL] got [", trim(access_mode), "]"
    stop 1
end if
    close(99)
end program test_inquire_file_property_matrix
