! vybe-test: fortran/io_list_directed_record_gaps/test_io_list_directed_record_gaps_parse_with_commas
! origin: languages/fortran/tests/fortran/test_io_list_directed_record_gaps.rs

program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a
    integer :: b
    text = '1, 2, 3'
    read(text, *) a
    read(text(4:80), *) b
    if ((a + b) /= 3) then
    print *, "FAIL: want [3] got [", a + b, "]"
    stop 1
end if
end program test_io_list_directed_record_gaps
