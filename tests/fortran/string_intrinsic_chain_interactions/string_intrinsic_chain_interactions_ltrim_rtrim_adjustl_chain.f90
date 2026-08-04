! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain
    character(len=12) :: text
    text = '   abc  '
    if ((len_trim(text)) /= 8) then
    print *, "FAIL: want [8] got [", len_trim(text), "]"
    stop 1
end if
    if ((len_trim(ltrim(text))) /= 3) then
    print *, "FAIL: want [3] got [", len_trim(ltrim(text)), "]"
    stop 1
end if
    if ((len_trim(rtrim(text))) /= 6) then
    print *, "FAIL: want [6] got [", len_trim(rtrim(text)), "]"
    stop 1
end if
    if ((len_trim(adjustl(text))) /= 3) then
    print *, "FAIL: want [3] got [", len_trim(adjustl(text)), "]"
    stop 1
end if
    if (trim(trim(adjustl(text))) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(adjustl(text)), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_ltrim_rtrim_adjustl_chain
