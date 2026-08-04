! vybe-test: fortran/named_loops/named_loop_with_call
! origin: languages/fortran/tests/fortran/test_named_loops.rs

program test
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 37 ]
    integer :: i, total
    total = 0
    accumulate: do i = 1, 10
        if (mod(i, 3) == 0) cycle accumulate
        call add(total, i)
    end do accumulate
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((total) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total, "]"
        stop 1
    end if
contains
    subroutine add(acc, n)
        integer, intent(inout) :: acc
        integer, intent(in)    :: n
        acc = acc + n
end subroutine add
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program test
