! vybe-test: fortran/fortran_procedure_quality/procedure_quality_host_association
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_host_association
    integer :: result
    integer :: base
    base = 4
    call caller(base, result)
    if ((result) /= 8) then
    print *, "FAIL: want [8] got [", result, "]"
    stop 1
end if

contains
    subroutine caller(x, out)
        integer, intent(in) :: x
        integer, intent(out) :: out
        out = double(x)
    end subroutine caller

    integer function double(v)
        integer, intent(in) :: v
        double = v * 2
    end function double
end program procedure_quality_host_association
