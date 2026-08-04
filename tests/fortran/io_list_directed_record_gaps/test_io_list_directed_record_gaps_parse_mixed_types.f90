! vybe-test: fortran/io_list_directed_record_gaps/test_io_list_directed_record_gaps_parse_mixed_types
! origin: languages/fortran/tests/fortran/test_io_list_directed_record_gaps.rs

program test_io_list_directed_record_gaps
    character(len=80) :: text
    integer :: a
    real :: x
    logical :: f
    text = '42 3.5 .true.'
    read(text, *) a, x, f
    if ((a) /= 42) then
    print *, "FAIL: want [42] got [", a, "]"
    stop 1
end if
    if ((int(x)) /= 3) then
    print *, "FAIL: want [3] got [", int(x), "]"
    stop 1
end if
    if ((f) .neqv. .true.) then
    print *, "FAIL: want [true] got [", f, "]"
    stop 1
end if
end program test_io_list_directed_record_gaps
