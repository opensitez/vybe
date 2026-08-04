! vybe-test: fortran/logical_eqv_neqv/three_way_eqv_chain_all_true_operands
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .eqv. .true. .eqv. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .eqv. .true. .eqv. .true., "]"
    stop 1
end if
end program t
