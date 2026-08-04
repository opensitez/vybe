! vybe-test: fortran/subroutine_extended/compile_recursive_function_with_result_clause
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs

program t
    if ((len_num(12345)) /= 5) then
    print *, "FAIL: want [5] got [", len_num(12345), "]"
    stop 1
end if
contains
    recursive function len_num(n) result(d)
        integer, intent(in) :: n
        integer :: d
        if (n < 10) then
            d = 1
        else
            d = 1 + len_num(n / 10)
        end if
    end function len_num
end program t
