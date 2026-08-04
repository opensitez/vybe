! vybe-test: fortran/logical_eqv_neqv/two_neqv_tf_results_are_eqv
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .neqv. .false.) .eqv. (.false. .neqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .neqv. .false.) .eqv. (.false. .neqv. .true.), "]"
    stop 1
end if
end program t
