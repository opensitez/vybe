! vybe-test: fortran/fortran2018/fortran_2018_out_of_range_runtime
! origin: languages/fortran/tests/fortran/test_fortran2018.rs

program t
    integer(kind=4) :: small = 100
    if ((out_of_range(small, 0_2))) then
    print *, "FAIL: want [0] got [", out_of_range(small, 0_2), "]"
    stop 1
end if
    if ((out_of_range(3.14, 0))) then
    print *, "FAIL: want [0] got [", out_of_range(3.14, 0), "]"
    stop 1
end if
end program t
