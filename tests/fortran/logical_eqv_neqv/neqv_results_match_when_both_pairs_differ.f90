! vybe-test: fortran/logical_eqv_neqv/neqv_results_match_when_both_pairs_differ
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .neqv. .false.) .eqv. (.false. .neqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .neqv. .false.) .eqv. (.false. .neqv. .true.), "]"
    stop 1
end if
end program t
