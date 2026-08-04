! vybe-test: fortran/do_construct_reentrancy/test_do_construct_reentrancy_bare_loop_with_exit
! origin: languages/fortran/tests/fortran/test_do_construct_reentrancy.rs

program test_do_construct_reentrancy_bare_loop
    integer :: i, total
    i = 0
    total = 0
    do
        i = i + 1
        if (i == 4) exit
        total = total + i
    end do
    if ((i) /= 4) then
    print *, "FAIL: want [4] got [", i, "]"
    stop 1
end if
    if ((total) /= 6) then
    print *, "FAIL: want [6] got [", total, "]"
    stop 1
end if
end program test_do_construct_reentrancy_bare_loop
