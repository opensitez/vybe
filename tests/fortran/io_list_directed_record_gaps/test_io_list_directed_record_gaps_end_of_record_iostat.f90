! vybe-test: fortran/io_list_directed_record_gaps/test_io_list_directed_record_gaps_end_of_record_iostat
! origin: languages/fortran/tests/fortran/test_io_list_directed_record_gaps.rs

program test_io_list_directed_record_gaps
    character(len=10) :: text
    integer :: x
    integer :: ios
    text = '7'
    read(text, *, iostat=ios) x
    read(text, *, iostat=ios) x
    if (ios /= 0) print *, 1
end program test_io_list_directed_record_gaps
