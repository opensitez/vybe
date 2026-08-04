! vybe-test: fortran/io_advanced/read_err_label_path
! origin: languages/fortran/tests/fortran/test_io_advanced.rs

program test
    integer :: n, ios
    character(len=12) :: buf
    n = -1
    read('bad', '(I0)', iostat=ios, err=100, end=200) n
    if ((n) /= -1) then
    print *, "FAIL: want [-1] got [", n, "]"
    stop 1
end if
    goto 300
100 print *, ios
    goto 300
200 print *, -1
300 continue
end program test
