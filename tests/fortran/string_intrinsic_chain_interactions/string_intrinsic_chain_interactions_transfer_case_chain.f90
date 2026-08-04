! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_transfer_case_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_transfer_case_chain
    character(len=18) :: source
    source = '  Mixed CASE Data  '
    if ((verify(adjustl(source), ' ', .false.)) /= 1) then
    print *, "FAIL: want [1] got [", verify(adjustl(source), ' ', .false.), "]"
    stop 1
end if
    if (trim(adjustl(transfer(source, ''))) /= "Mixed CASE Data") then
    print *, "FAIL: want [Mixed CASE Data] got [", adjustl(transfer(source, '')), "]"
    stop 1
end if
    if ((len_trim(transfer(trim(adjustl(source)), ''))) /= 16) then
    print *, "FAIL: want [16] got [", len_trim(transfer(trim(adjustl(source)), '')), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_transfer_case_chain
