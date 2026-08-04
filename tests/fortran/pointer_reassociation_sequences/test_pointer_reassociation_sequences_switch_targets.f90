! vybe-test: fortran/pointer_reassociation_sequences/test_pointer_reassociation_sequences_switch_targets
! origin: languages/fortran/tests/fortran/test_pointer_reassociation_sequences.rs

program test_pointer_reassociation_sequences
    integer, target :: a
    integer, target :: b
    integer, pointer :: p

    a = 5
    b = 9
    p => a
    if ((p) /= 5) then
    print *, "FAIL: want [5] got [", p, "]"
    stop 1
end if
    p => b
    if ((p) /= 9) then
    print *, "FAIL: want [9] got [", p, "]"
    stop 1
end if
end program test_pointer_reassociation_sequences
