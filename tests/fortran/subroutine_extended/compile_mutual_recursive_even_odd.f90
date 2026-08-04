! vybe-test: fortran/subroutine_extended/compile_mutual_recursive_even_odd
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    if ((is_even(6)) .neqv. .true.) then
    print *, "FAIL: want [true] got [", is_even(6), "]"
    stop 1
end if
contains
    recursive function is_even(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .true.
        else
            b = is_odd(n - 1)
        end if
    end function is_even

    recursive function is_odd(n) result(b)
        integer, intent(in) :: n
        logical :: b
        if (n == 0) then
            b = .false.
        else
            b = is_even(n - 1)
        end if
    end function is_odd
end program t
