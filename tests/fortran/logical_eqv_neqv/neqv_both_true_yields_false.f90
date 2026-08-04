! vybe-test: fortran/logical_eqv_neqv/neqv_both_true_yields_false
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .neqv. .true.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .true. .neqv. .true., "]"
    stop 1
end if
end program t
