! vybe-test: fortran/logical_operators_and_parentheses/test_logical_eqv_neqv_precedence
! origin: languages/fortran/tests/fortran/test_logical_operators_and_parentheses.rs

program test_logical_eqv_neqv_precedence
    if ((.true. .eqv. .true.) .neqv. .true.) then
    print *, "FAIL: want [True] got [", .true. .eqv. .true., "]"
    stop 1
end if
    if ((.true. .neqv. .true.) .neqv. .false.) then
    print *, "FAIL: want [False] got [", .true. .neqv. .true., "]"
    stop 1
end if
    if ((.true. .or. (.false. .eqv. .true.)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", .true. .or. (.false. .eqv. .true.), "]"
    stop 1
end if
end program test_logical_eqv_neqv_precedence
