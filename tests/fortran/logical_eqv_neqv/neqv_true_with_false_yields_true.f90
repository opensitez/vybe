! vybe-test: fortran/logical_eqv_neqv/neqv_true_with_false_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .neqv. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .neqv. .false., "]"
    stop 1
end if
end program t
