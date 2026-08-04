! vybe-test: fortran/print_io/print_logical_true
! origin: languages/fortran/tests/fortran/test_print_io.rs
program t
if ((.true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true., "]"
    stop 1
end if
end program t
