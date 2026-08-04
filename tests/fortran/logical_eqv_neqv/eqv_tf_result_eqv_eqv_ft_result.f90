! vybe-test: fortran/logical_eqv_neqv/eqv_tf_result_eqv_eqv_ft_result
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .false.) .eqv. (.false. .eqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .eqv. .false.) .eqv. (.false. .eqv. .true.), "]"
    stop 1
end if
end program t
