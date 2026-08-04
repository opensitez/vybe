! vybe-test: fortran/pure_elemental/optional_integer
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, with_default(5)
    print *, with_default(5, 10)
contains
    function with_default(x, y) result(res)
        integer, intent(in) :: x
        integer, intent(in), optional :: y
        integer :: res
        if (present(y)) then
            res = x + y
        else
            res = x
        end if
    end function with_default
end program test
