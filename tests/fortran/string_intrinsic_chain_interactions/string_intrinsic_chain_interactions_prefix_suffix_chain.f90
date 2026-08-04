! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_prefix_suffix_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_prefix_suffix_chain
    character(len=12) :: text
    character(len=8) :: head
    text = '  fortran  '
    head = trim(adjustl(text))
    if ((len_trim(head)) /= 7) then
    print *, "FAIL: want [7] got [", len_trim(head), "]"
    stop 1
end if
    if (trim(head // '_ok') /= "fortran_ok") then
    print *, "FAIL: want [fortran_ok] got [", head // '_ok', "]"
    stop 1
end if
    if ((index(head, 'tran')) /= 4) then
    print *, "FAIL: want [4] got [", index(head, 'tran'), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_prefix_suffix_chain
