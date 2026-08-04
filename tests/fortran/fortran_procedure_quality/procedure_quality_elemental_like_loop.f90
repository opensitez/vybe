! vybe-test: fortran/fortran_procedure_quality/procedure_quality_elemental_like_loop
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_elemental_like_loop
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 120 ]
    integer :: i
    integer :: out
    out = 1
    do i = 1, 5
        call set_scale(i, out)
    end do
        vybe_check_i = vybe_check_i + 1
    if (vybe_check_i > 1) then
        print *, "FAIL: more than 1 line(s)"
        stop 1
    end if
    if ((out) /= vybe_check_w(vybe_check_i)) then
        print *, "FAIL at ", vybe_check_i, " got [", out, "]"
        stop 1
    end if

contains
    subroutine set_scale(v, state)
        integer, intent(in) :: v
        integer, intent(inout) :: state
        state = state * v
    end subroutine set_scale
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program procedure_quality_elemental_like_loop
