! vybe-test: fortran/pure_elemental/recursive_ackermann
! origin: languages/fortran/tests/fortran/test_pure_elemental.rs

program test
    print *, ack(2, 3)
contains
    recursive function ack(m, n) result(res)
        integer, intent(in) :: m, n
        integer :: res
        if (m == 0) then
            res = n + 1
        else if (n == 0) then
            res = ack(m - 1, 1)
        else
            res = ack(m - 1, ack(m, n - 1))
        end if
    end function ack
end program test
