! vybe-test: fortran/pure_elemental/recursive_gcd
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, gcd(48, 18)
contains
    recursive function gcd(a, b) result(res)
        integer, intent(in) :: a, b
        integer :: res
        if (b == 0) then
            res = a
        else
            res = gcd(b, mod(a, b))
        end if
    end function gcd
end program test
