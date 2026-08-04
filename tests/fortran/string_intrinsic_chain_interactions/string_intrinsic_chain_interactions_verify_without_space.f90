! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_verify_without_space
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_verify_without_space
    character(len=12) :: text
    text = 'abc123def456'
    if ((verify(text, '0123456789', .true.)) /= 1) then
    print *, "FAIL: want [1] got [", verify(text, '0123456789', .true.), "]"
    stop 1
end if
    if ((scan(text, '123')) /= 4) then
    print *, "FAIL: want [4] got [", scan(text, '123'), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_verify_without_space
