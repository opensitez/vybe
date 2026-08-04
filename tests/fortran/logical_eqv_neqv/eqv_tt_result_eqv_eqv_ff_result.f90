! vybe-test: fortran/logical_eqv_neqv/eqv_tt_result_eqv_eqv_ff_result
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .true.) .eqv. (.false. .eqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .eqv. .true.) .eqv. (.false. .eqv. .false.), "]"
    stop 1
end if
end program t
