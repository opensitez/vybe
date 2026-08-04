! vybe-test: fortran/string_intrinsic_chain_interactions/string_intrinsic_chain_interactions_char_comparisons
! origin: languages/fortran/tests/fortran/test_string_intrinsic_chain_interactions.rs

program string_intrinsic_chain_interactions_char_comparisons
    character(len=8) :: left
    character(len=8) :: right
    left = 'alpha'
    right = 'alphA'
    if ((left < right) .neqv. .false.) then
    print *, "FAIL: want [False] got [", left < right, "]"
    stop 1
end if
    if (trim(trim(merge(left, right, left > right))) /= "alphA") then
    print *, "FAIL: want [alphA] got [", trim(merge(left, right, left > right)), "]"
    stop 1
end if
end program string_intrinsic_chain_interactions_char_comparisons
