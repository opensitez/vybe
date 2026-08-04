! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_factorial_guard
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_factorial_guard
    if ((fact(0)) /= 1) then
    print *, "FAIL: want [1] got [", fact(0), "]"
    stop 1
end if
    if ((fact(5)) /= 120) then
    print *, "FAIL: want [120] got [", fact(5), "]"
    stop 1
end if
contains
    recursive integer function fact(n) result(out)
        integer, intent(in) :: n
        if (n <= 1) then
            out = 1
        else
            out = n * fact(n - 1)
        end if
    end function fact
end program recursive_procedures_and_terminators_factorial_guard
