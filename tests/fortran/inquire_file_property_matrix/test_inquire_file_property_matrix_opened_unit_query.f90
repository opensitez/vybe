! vybe-test: fortran/inquire_file_property_matrix/test_inquire_file_property_matrix_opened_unit_query
! origin: languages/fortran/tests/fortran/test_inquire_file_property_matrix.rs

program test_inquire_file_property_matrix
    integer :: unit
    logical :: is_open
    open(newunit=unit, file='tmp_inquire_opened_probe.txt', status='replace')
    inquire(unit=unit, opened=is_open)
    if (trim(is_open) /= ".TRUE.") then
    print *, "FAIL: want [.TRUE.] got [", is_open, "]"
    stop 1
end if
    close(unit)
    if ((0) /= 0) then
    print *, "FAIL: want [0] got [", 0, "]"
    stop 1
end if
end program test_inquire_file_property_matrix
