! vybe-test: fortran/io_pure_output_buffering/test_io_pure_output_buffering_prints_as_expected
! origin: languages/fortran/tests/fortran/test_io_pure_output_buffering.rs

program test_io_pure_output_buffering
    if ((1) /= 1) then
    print *, "FAIL: want [1] got [", 1, "]"
    stop 1
end if
    if ((2) /= 2) then
    print *, "FAIL: want [2] got [", 2, "]"
    stop 1
end if
    if ((3) /= 3) then
    print *, "FAIL: want [3] got [", 3, "]"
    stop 1
end if
end program test_io_pure_output_buffering
