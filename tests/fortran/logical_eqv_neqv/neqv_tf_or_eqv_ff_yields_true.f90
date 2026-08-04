! vybe-test: fortran/logical_eqv_neqv/neqv_tf_or_eqv_ff_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .neqv. .false.) .or. (.false. .eqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .neqv. .false.) .or. (.false. .eqv. .false.), "]"
    stop 1
end if
end program t
