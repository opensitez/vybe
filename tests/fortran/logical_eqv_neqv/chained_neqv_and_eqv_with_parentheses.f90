! vybe-test: fortran/logical_eqv_neqv/chained_neqv_and_eqv_with_parentheses
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((((.true. .neqv. .false.) .and. (.false. .eqv. .true.)) .eqv. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", ((.true. .neqv. .false.) .and. (.false. .eqv. .true.)) .eqv. .true., "]"
    stop 1
end if
end program t
