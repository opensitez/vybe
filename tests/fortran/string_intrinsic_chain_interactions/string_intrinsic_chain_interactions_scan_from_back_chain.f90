! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_scan_from_back_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_scan_from_back_chain
    character(len=16) :: text
    text = '  one, two, three  '
    if (trim(trim(adjustl(trim(text)))) /= "one, two, three") then
    print *, "FAIL: want [one, two, three] got [", trim(adjustl(trim(text))), "]"
    stop 1
end if
    if ((scan(trim(adjustl(trim(text))), ' ,', .true.)) /= 10) then
    print *, "FAIL: want [10] got [", scan(trim(adjustl(trim(text))), ' ,', .true.), "]"
    stop 1
end if
    if ((index(text, 'three')) /= 13) then
    print *, "FAIL: want [13] got [", index(text, 'three'), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_scan_from_back_chain
