! vybe-test: fortran/pure_elemental/recursive_power
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, power(2, 8)
contains
    recursive function power(base, exp) result(res)
        integer, intent(in) :: base, exp
        integer :: res
        if (exp == 0) then
            res = 1
        else
            res = base * power(base, exp - 1)
        end if
    end function power
end program test
