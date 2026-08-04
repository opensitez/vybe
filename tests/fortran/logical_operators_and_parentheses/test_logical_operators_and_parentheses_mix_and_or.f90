! vybe-test: fortran/logical_operators_and_parentheses/test_logical_operators_and_parentheses_mix_and_or
! origin: languages/fortran/tests/fortran/test_logical_operators_and_parentheses.rs

program test_logical_operators_and_parentheses
    logical :: a
    logical :: b
    logical :: c
    a = .true.
    b = .false.
    c = .true.
    if (((a .and. .not. b) .or. c) .neqv. .true.) then
    print *, "FAIL: want [True] got [", (a .and. .not. b) .or. c, "]"
    stop 1
end if
    if (((a .and. (b .or. c))) .neqv. .true.) then
    print *, "FAIL: want [True] got [", (a .and. (b .or. c)), "]"
    stop 1
end if
end program test_logical_operators_and_parentheses
