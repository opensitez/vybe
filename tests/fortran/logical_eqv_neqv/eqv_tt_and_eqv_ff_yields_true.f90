! vybe-test: fortran/logical_eqv_neqv/eqv_tt_and_eqv_ff_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .true.) .and. (.false. .eqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .eqv. .true.) .and. (.false. .eqv. .false.), "]"
    stop 1
end if
end program t
