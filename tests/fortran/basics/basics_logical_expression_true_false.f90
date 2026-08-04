! vybe-test: fortran/basics/basics_logical_expression_true_false
! origin: languages/fortran/tests/fortran/test_basics.rs

program test
    implicit none
    integer :: x
    logical :: is_big
    x = 7
    is_big = (x > 5)
    if ((is_big) .neqv. .true.) then
    print *, "FAIL: want [true] got [", is_big, "]"
    stop 1
end if
    if ((.not. is_big) .neqv. .false.) then
    print *, "FAIL: want [false] got [", .not. is_big, "]"
    stop 1
end if
end program test
