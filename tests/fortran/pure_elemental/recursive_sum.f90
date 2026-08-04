! vybe-test: fortran/pure_elemental/recursive_sum
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, rsum(10)
contains
    recursive function rsum(n) result(res)
        integer, intent(in) :: n
        integer :: res
        if (n <= 0) then
            res = 0
        else
            res = n + rsum(n - 1)
        end if
    end function rsum
end program test
