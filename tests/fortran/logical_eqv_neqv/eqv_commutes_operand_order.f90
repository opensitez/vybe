! vybe-test: fortran/logical_eqv_neqv/eqv_commutes_operand_order
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.false. .eqv. .true.) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .false. .eqv. .true., "]"
    stop 1
end if
end program t
