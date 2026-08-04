! vybe-test: fortran/do_construct_reentrancy/test_do_construct_reentrancy_descending_inner_loop
! origin: languages/fortran/tests/fortran/test_do_construct_reentrancy.rs

program test_do_construct_reentrancy_descending
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 18 ]
    integer :: outer, inner, total
    total = 0
    do outer = 1, 2
        do inner = 5, 1, -2
            total = total + inner
        end do
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test_do_construct_reentrancy_descending
