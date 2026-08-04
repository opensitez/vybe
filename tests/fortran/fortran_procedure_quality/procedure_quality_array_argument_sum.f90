! vybe-test: fortran/fortran_procedure_quality/procedure_quality_array_argument_sum
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_array_argument_sum
integer :: vybe_check_i = 0
integer :: vybe_check_w(1) = [ 15 ]
    integer, dimension(5) :: values
    integer :: total
    values = (/ 1, 2, 3, 4, 5 /)
    call array_sum(values, total)
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
    subroutine array_sum(v, total)
        integer, intent(in) :: v(:)
        integer, intent(out) :: total
        integer :: i
        total = 0
        do i = 1, size(v)
            total = total + v(i)
        end do
    end subroutine array_sum
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program procedure_quality_array_argument_sum
