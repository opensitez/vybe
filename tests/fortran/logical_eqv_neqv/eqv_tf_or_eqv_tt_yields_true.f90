! vybe-test: fortran/logical_eqv_neqv/eqv_tf_or_eqv_tt_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .false.) .or. (.true. .eqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .eqv. .false.) .or. (.true. .eqv. .true.), "]"
    stop 1
end if
end program t
