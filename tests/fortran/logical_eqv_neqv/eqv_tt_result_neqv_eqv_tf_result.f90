! vybe-test: fortran/logical_eqv_neqv/eqv_tt_result_neqv_eqv_tf_result
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .true.) .neqv. (.true. .eqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .eqv. .true.) .neqv. (.true. .eqv. .false.), "]"
    stop 1
end if
end program t
