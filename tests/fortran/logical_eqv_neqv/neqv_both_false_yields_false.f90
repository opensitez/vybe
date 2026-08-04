! vybe-test: fortran/logical_eqv_neqv/neqv_both_false_yields_false
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.false. .neqv. .false.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .false. .neqv. .false., "]"
    stop 1
end if
end program t
