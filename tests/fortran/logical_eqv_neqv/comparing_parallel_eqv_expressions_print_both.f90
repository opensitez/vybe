! vybe-test: fortran/logical_eqv_neqv/comparing_parallel_eqv_expressions_print_both
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .eqv. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .eqv. .true., "]"
    stop 1
end if
if ((.false. .eqv. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .false. .eqv. .false., "]"
    stop 1
end if
end program t
