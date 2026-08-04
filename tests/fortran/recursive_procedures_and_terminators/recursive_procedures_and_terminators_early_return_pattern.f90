! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_early_return_pattern
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_early_return_pattern
    if ((first_nonzero((/0, 0, 5, 7/), 1)) /= 5) then
    print *, "FAIL: want [5] got [", first_nonzero((/0, 0, 5, 7/), 1), "]"
    stop 1
end if
contains
    recursive integer function first_nonzero(values, idx) result(out)
        integer, intent(in) :: values(:)
        integer, intent(in) :: idx
        if (idx > size(values)) then
            out = -1
        else if (values(idx) /= 0) then
            out = values(idx)
        else
            out = first_nonzero(values, idx + 1)
        end if
    end function first_nonzero
end program recursive_procedures_and_terminators_early_return_pattern
