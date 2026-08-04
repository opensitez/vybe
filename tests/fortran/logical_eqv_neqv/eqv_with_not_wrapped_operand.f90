! vybe-test: fortran/logical_eqv_neqv/eqv_with_not_wrapped_operand
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.true. .eqv. .not. .false.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .true. .eqv. .not. .false., "]"
    stop 1
end if
end program t
