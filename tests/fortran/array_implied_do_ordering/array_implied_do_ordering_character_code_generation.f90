! vybe-test: fortran/array_implied_do_ordering/array_implied_do_ordering_character_code_generation
! origin: languages/fortran/tests/fortran/test_array_implied_do_ordering.rs

program array_implied_do_ordering_character_code_generation
    character(len=1), allocatable :: values(:)
    values = (/ (achar(iachar('a') + i), i = 0, 3) /)
    if ((size(values)) /= 4) then
    print *, "FAIL: want [4] got [", size(values), "]"
    stop 1
end if
    if ((len(values(1))) /= 1) then
    print *, "FAIL: want [1] got [", len(values(1)), "]"
    stop 1
end if
    if ((iachar(values(1))) /= 97) then
    print *, "FAIL: want [97] got [", iachar(values(1)), "]"
    stop 1
end if
    if ((iachar(values(4))) /= 100) then
    print *, "FAIL: want [100] got [", iachar(values(4)), "]"
    stop 1
end if
end program array_implied_do_ordering_character_code_generation
