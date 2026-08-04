! vybe-test: fortran/logical_eqv_neqv/compare_two_eqv_expressions_with_neqv
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .eqv. .false.) .neqv. (.false. .eqv. .true.)) .neqv. .false.) then
    print *, "FAIL: want [false] got [", (.true. .eqv. .false.) .neqv. (.false. .eqv. .true.), "]"
    stop 1
end if
end program t
