! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_index_scan_verify_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_index_scan_verify_chain
    character(len=16) :: x
    x = '  one,two,three  '
    if ((index(trim(x), ',')) /= 4) then
    print *, "FAIL: want [4] got [", index(trim(x), ','), "]"
    stop 1
end if
    if ((scan(trim(x), ',')) /= 4) then
    print *, "FAIL: want [4] got [", scan(trim(x), ','), "]"
    stop 1
end if
    if ((verify(trim(x), '1234567890, ', .true.)) /= 1) then
    print *, "FAIL: want [1] got [", verify(trim(x), '1234567890, ', .true.), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_index_scan_verify_chain
