! vybe-test: fortran/interfaces/if_recursive_result_11
! origin: languages/fortran/tests/fortran/test_interfaces.rs
program t
integer :: v
v = f(5)
if (v /= 120) then
    print *, "FAIL: want [120] got [", v, "]"
    stop 1
end if
contains
recursive integer function f(n) result(r)
integer::n
if (n <= 1) then
    r = 1
else
    r = n * f(n - 1)
end if
end function f
end program t
