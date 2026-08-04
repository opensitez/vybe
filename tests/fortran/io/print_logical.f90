! vybe-test: fortran/io/print_logical
! origin: languages/fortran/tests/fortran/test_io.rs

program test
    if ((.true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true., "]"
    stop 1
end if
    if ((.false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .false., "]"
    stop 1
end if
end program test
