! vybe-test: fortran/logical_eqv_neqv/eqv_both_false_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.false. .eqv. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .false. .eqv. .false., "]"
    stop 1
end if
end program t
