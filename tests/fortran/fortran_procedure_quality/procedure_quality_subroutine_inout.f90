! vybe-test: fortran/fortran_procedure_quality/procedure_quality_subroutine_inout
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_subroutine_inout
    integer :: left
    integer :: right
    integer :: output
    left = 4
    right = 5
    call add_pair(left, right, output)
    if ((output) /= 9) then
    print *, "FAIL: want [9] got [", output, "]"
    stop 1
end if

contains
    subroutine add_pair(a, b, c)
        integer, intent(in) :: a
        integer, intent(in) :: b
        integer, intent(out) :: c
        c = a + b
    end subroutine add_pair
end program procedure_quality_subroutine_inout
