! vybe-test: fortran/do_construct_step_semantics/test_do_construct_step_semantics_single_iteration_from_expression_bounds
! origin: languages/fortran/tests/fortran/test_do_construct_step_semantics.rs

program test_do_construct_step_semantics_single_expression_iteration
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 2 ]
    integer :: i
    integer :: total
    total = 0
    do i = 2 + 0, 4 - 2, 1 + 1
        total = total + i
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
end program test_do_construct_step_semantics_single_expression_iteration
