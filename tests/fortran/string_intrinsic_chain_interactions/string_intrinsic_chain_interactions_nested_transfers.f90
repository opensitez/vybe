! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_nested_transfers
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_nested_transfers
    character(len=4) :: x
    character(len=12) :: y
    x = 'ab'
    y = trim(adjustl(x // repeat('c', 2)))
    if (trim(y) /= "abcc") then
    print *, "FAIL: want [abcc] got [", y, "]"
    stop 1
end if
    if ((len_trim(y)) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(y), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_nested_transfers
