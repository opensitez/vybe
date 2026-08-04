! vybe-test: fortran/logical_eqv_neqv/neqv_or_eqv_mixed_with_parentheses
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .neqv. .false.) .or. (.true. .eqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .neqv. .false.) .or. (.true. .eqv. .false.), "]"
    stop 1
end if
end program t
