! vybe-test: fortran/fortran_io_quality/io_quality_logical_output
! origin: languages/fortran/tests/fortran/test_fortran_io_quality.rs

program io_quality_logical_output
    logical :: enabled
    enabled = .true.
    if ((enabled) .neqv. .true.) then
    print *, "FAIL: want [true] got [", enabled, "]"
    stop 1
end if
end program io_quality_logical_output
