! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_adjustl_adjustr
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_adjustl_adjustr
    character(len=10) :: text
    text = '  mix ' 
    if (trim(trim(adjustl(text))) /= "mix") then
    print *, "FAIL: want [mix] got [", trim(adjustl(text)), "]"
    stop 1
end if
    if ((len_trim(adjustr(text))) /= 4) then
    print *, "FAIL: want [4] got [", len_trim(adjustr(text)), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_adjustl_adjustr
