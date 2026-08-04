! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_trim_len_index
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_trim_len_index
    character(len=20) :: raw
    raw = '  vybe-fortran  '
    if ((len(raw)) /= 20) then
    print *, "FAIL: want [20] got [", len(raw), "]"
    stop 1
end if
    if ((len_trim(raw)) /= 12) then
    print *, "FAIL: want [12] got [", len_trim(raw), "]"
    stop 1
end if
    if ((index(trim(raw), 'fortran')) /= 5) then
    print *, "FAIL: want [5] got [", index(trim(raw), 'fortran'), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_trim_len_index
