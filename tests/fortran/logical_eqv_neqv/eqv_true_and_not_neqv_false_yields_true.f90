! vybe-test: fortran/logical_eqv_neqv/eqv_true_and_not_neqv_false_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .true.) .and. .not. (.true. .neqv. .false.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", (.true. .eqv. .true.) .and. .not. (.true. .neqv. .false.), "]"
    stop 1
end if
end program t
