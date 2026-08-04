! vybe-test: fortran/logical_eqv_neqv/eqv_case_insensitive_logical_literals
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.TRUE. .and. .FALSE.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .TRUE. .and. .FALSE., "]"
    stop 1
end if
if ((.true. .or. .FALSE.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .or. .FALSE., "]"
    stop 1
end if
end program t
