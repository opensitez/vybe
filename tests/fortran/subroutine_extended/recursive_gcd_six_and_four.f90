! vybe-test: fortran/subroutine_extended/recursive_gcd_six_and_four
! origin: languages/fortran/tests/fortran/test_subroutine_extended.rs
program t
if ((my_gcd(6, 4)) /= 2) then
    print *, "FAIL: want [2] got [", my_gcd(6, 4), "]"
    stop 1
end if
contains
recursive function my_gcd(a, b) result(g)
integer, intent(in) :: a, b
integer :: g
if (b == 0) then
g = a
else
g = my_gcd(b, mod(a, b))
end if
end function my_gcd
end program t
