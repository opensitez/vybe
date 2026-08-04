! vybe-test: fortran/full_programs/full_program_internal_procedures
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program full_program_internal_procedures
integer :: vybe_check_i = 0
integer :: vybe_check_w(2) = [ 10, 28 ]
    integer :: x
    x = 7
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((adjust(x)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", adjust(x), "]"
        stop 1
    end if
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 2) then
        print *, "FAIL: more than 2 line(s)"
        stop 1
    end if
    if ((total(x)) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", total(x), "]"
        stop 1
    end if
contains
    integer function adjust(v)
        integer, intent(in) :: v
        adjust = v + 3
    end function adjust

    integer function total(v)
        integer, intent(in) :: v
        integer :: i
        total = 0
        do i = 1, v
            total = total + i
        end do
    end function total
if (vybe_check_i /= 2) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 2"
    stop 1
end if
end program full_program_internal_procedures
