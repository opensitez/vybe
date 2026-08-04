! vybe-test: fortran/else_if_cascade_priority/test_else_if_cascade_truth_table_ordering
! origin: languages/fortran/tests/fortran/test_else_if_cascade_priority.rs

program test_else_if_cascade_priority_truth
    logical :: a
    logical :: b
    integer :: c
    a = .true.
    b = .false.
    c = 0
    if (.not. (a .and. b)) then
        c = 10
    else if (a .and. .not. b) then
        c = 20
    else if (b) then
        c = 30
    end if
    if ((c) /= 10) then
    print *, "FAIL: want [10] got [", c, "]"
    stop 1
end if
end program test_else_if_cascade_priority_truth
