! vybe-test: fortran/logical_eqv_neqv/eqv_results_differ_when_one_side_flips
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .true.) .eqv. (.true. .eqv. .false.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", (.true. .eqv. .true.) .eqv. (.true. .eqv. .false.), "]"
    stop 1
end if
end program t
