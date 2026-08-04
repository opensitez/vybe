! vybe-test: fortran/logical_eqv_neqv/neqv_with_not_wrapped_operand
! origin: languages/fortran/tests/fortran/test_logical_eqv_neqv.rs
program t
if ((.false. .neqv. .not. .true.) .neqv. .true.) then
    print *, "FAIL: want [true] got [", .false. .neqv. .not. .true., "]"
    stop 1
end if
end program t
