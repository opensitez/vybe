! vybe-test: fortran/logical_eqv_neqv/eqv_ff_or_neqv_tt_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.false. .eqv. .false.) .or. (.true. .neqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.false. .eqv. .false.) .or. (.true. .neqv. .true.), "]"
    stop 1
end if
end program t
