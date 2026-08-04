! vybe-test: fortran/logical_eqv_neqv/neqv_false_with_true_yields_true
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.false. .neqv. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .false. .neqv. .true., "]"
    stop 1
end if
end program t
