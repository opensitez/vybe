! vybe-test: fortran/recursive_procedures_and_terminators/recursive_procedures_and_terminators_mutual_recursion_odd_even
! origin: languages/fortran/tests/fortran/test_recursive_procedures_and_terminators.rs

program recursive_procedures_and_terminators_mutual_recursion_odd_even
    if ((is_even(4)) .neqv. .true.) then
    print *, "FAIL: want [True] got [", is_even(4), "]"
    stop 1
end if
    if ((is_odd(4)) .neqv. .false.) then
    print *, "FAIL: want [False] got [", is_odd(4), "]"
    stop 1
end if
contains
    recursive logical function is_even(n)
        integer, intent(in) :: n
        if (n == 0) then
            is_even = .true.
        else
            is_even = is_odd(n - 1)
        end if
    end function is_even

    recursive logical function is_odd(n)
        integer, intent(in) :: n
        if (n == 0) then
            is_odd = .false.
        else
            is_odd = is_even(n - 1)
        end if
    end function is_odd
end program recursive_procedures_and_terminators_mutual_recursion_odd_even
