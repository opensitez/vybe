! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_repeat_concat_and_len
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_repeat_concat_and_len
    character(len=20) :: text
    text = trim('A') // trim(repeat('B', 3))
    if (trim(text) /= "ABBB") then
    print *, "FAIL: want [ABBB] got [", text, "]"
    stop 1
end if
    if ((len_trim(text)) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(text), "]"
    stop 1
end if
    if ((index(text, 'BBB')) /= 2) then
    print *, "FAIL: want [2] got [", index(text, 'BBB'), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_repeat_concat_and_len
