! vybe-test: fortran/io_list_directed_record_gaps/test_io_list_directed_record_gaps_handles_multiple_whitespace
! origin: languages/fortran/tests/fortran/test_io_list_directed_record_gaps.rs

program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a, b, c
    text = '  10    20   30  '
    read(text, *) a, b, c
    if ((a + b + c) /= 60) then
    print *, "FAIL: want [60] got [", a + b + c, "]"
    stop 1
end if
end program test_io_list_directed_record_gaps
