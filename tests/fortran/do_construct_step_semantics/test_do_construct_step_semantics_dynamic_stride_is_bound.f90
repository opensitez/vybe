! vybe-test: fortran/do_construct_step_semantics/test_do_construct_step_semantics_dynamic_stride_is_bound
! origin: languages/fortran/tests/fortran/test_do_construct_step_semantics.rs

program test_do_construct_step_semantics
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 50 ]
    integer :: i
    integer :: step
    integer :: total
    total = 0
    step = 4
    do i = 2, 20, step
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
end program test_do_construct_step_semantics
