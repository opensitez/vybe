! vybe-test: fortran/pointer_alloc_extended/pointer_chain_across_two_targets
! origin: languages/fortran/tests/fortran/test_pointer_alloc_extended.rs
program t
integer, target :: left = 5, right = 15
integer, pointer :: hop
hop => left
if ((hop) /= 5) then
    print *, "FAIL: want [5] got [", hop, "]"
    stop 1
end if
hop => right
if ((hop) /= 15) then
    print *, "FAIL: want [15] got [", hop, "]"
    stop 1
end if
end program t
