! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_merge_adjustl_chain
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_merge_adjustl_chain
    if (trim(trim(merge(adjustl('  x  '), adjustr('  y  '), len_trim('  x  ') > 0))) /= "x") then
    print *, "FAIL: want [x] got [", trim(merge(adjustl('  x  '), adjustr('  y  '), len_trim('  x  ') > 0)), "]"
    stop 1
end if
    if (trim(trim(merge(adjustl('  x  '), adjustr('  y  '), .false.))) /= "y") then
    print *, "FAIL: want [y] got [", trim(merge(adjustl('  x  '), adjustr('  y  '), .false.)), "]"
    stop 1
end if
    if ((len_trim(adjustl(merge('  xx  ', '  yy  ', .false.)))) /= 2) then
    print *, "FAIL: want [2] got [", len_trim(adjustl(merge('  xx  ', '  yy  ', .false.))), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_merge_adjustl_chain
