! vybe-test: fortran/logical_eqv_neqv/eqv_true_with_false_yields_false
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .eqv. .false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .true. .eqv. .false., "]"
    stop 1
end if
end program t
