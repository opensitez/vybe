! vybe-test: fortran/logical_eqv_neqv/compare_two_neqv_expressions_with_eqv
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if (((.true. .neqv. .true.) .eqv. (.false. .neqv. .false.)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", (.true. .neqv. .true.) .eqv. (.false. .neqv. .false.), "]"
    stop 1
end if
end program t
