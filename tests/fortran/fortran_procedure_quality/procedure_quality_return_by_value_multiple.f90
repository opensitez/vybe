! vybe-test: fortran/fortran_procedure_quality/procedure_quality_return_by_value_multiple
! origin: languages/fortran/tests/fortran/test_fortran_procedure_quality.rs

program procedure_quality_return_by_value_multiple
    integer :: first
    integer :: second
    call minmax(8, 3, first, second)
    if ((first) /= 3) then
    print *, "FAIL: want [3] got [", first, "]"
    stop 1
end if
    if ((second) /= 8) then
    print *, "FAIL: want [8] got [", second, "]"
    stop 1
end if

contains
    subroutine minmax(a, b, lo, hi)
        integer, intent(in) :: a
        integer, intent(in) :: b
        integer, intent(out) :: lo
        integer, intent(out) :: hi
        if (a < b) then
            lo = a
            hi = b
        else
            lo = b
            hi = a
        end if
    end subroutine minmax
end program procedure_quality_return_by_value_multiple
