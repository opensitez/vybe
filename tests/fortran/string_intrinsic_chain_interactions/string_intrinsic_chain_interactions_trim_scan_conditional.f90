! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_trim_scan_conditional
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_trim_scan_conditional
    character(len=20) :: base
    base = 'alpha;beta;gamma'
    if ((index(trim(base), ';')) /= 6) then
    print *, "FAIL: want [6] got [", index(trim(base), ';'), "]"
    stop 1
end if
    if ((scan(trim(base), ';', .false.)) /= 6) then
    print *, "FAIL: want [6] got [", scan(trim(base), ';', .false.), "]"
    stop 1
end if
    if ((verify(merge(base, 'fallback', len(base) > 10), 'abcdefghijklmnopqrstuvwxyz', .false.)) /= 6) then
    print *, "FAIL: want [6] got [", verify(merge(base, 'fallback', len(base) > 10), 'abcdefghijklmnopqrstuvwxyz', .false.), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_trim_scan_conditional
