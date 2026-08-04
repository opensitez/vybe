! vybe-test: fortran/logical_operators_and_parentheses/test_not_with_parentheses
! origin: languages/fortran/tests/fortran/test_logical_operators_and_parentheses.rs

program test_not_with_parentheses
    if ((.not. (.true. .and. .false.)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", .not. (.true. .and. .false.), "]"
    stop 1
end if
    if ((.not. (.false. .or. .false.)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", .not. (.false. .or. .false.), "]"
    stop 1
end if
    if (((.not. .true.) .and. .true.) .neqv. .false.) then
    print *, "FAIL: want [False] got [", (.not. .true.) .and. .true., "]"
    stop 1
end if
end program test_not_with_parentheses
