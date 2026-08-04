! vybe-test: fortran/print_io/print_logical_false
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((.false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .false., "]"
    stop 1
end if
end program t
