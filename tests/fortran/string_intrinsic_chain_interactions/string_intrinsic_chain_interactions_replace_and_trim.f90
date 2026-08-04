! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_replace_and_trim
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_replace_and_trim
    character(len=20) :: text
    text = 'the quick brown fox'
    if (trim(trim(adjustl(replace(text, 'quick', 'fast')))) /= "the fast brown fox") then
    print *, "FAIL: want [the fast brown fox] got [", trim(adjustl(replace(text, 'quick', 'fast'))), "]"
    stop 1
end if
    if ((len_trim(adjustl(replace(text, 'quick', 'fast')))) /= 17) then
    print *, "FAIL: want [17] got [", len_trim(adjustl(replace(text, 'quick', 'fast'))), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_replace_and_trim
