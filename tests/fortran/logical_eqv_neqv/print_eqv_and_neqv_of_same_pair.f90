! vybe-test: fortran/logical_eqv_neqv/print_eqv_and_neqv_of_same_pair
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .eqv. .false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .true. .eqv. .false., "]"
    stop 1
end if
if ((.true. .neqv. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .neqv. .false., "]"
    stop 1
end if
end program t
