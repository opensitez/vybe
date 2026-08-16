! vybe-test: fortran/recursive_results/recursive_results_09
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if ((f(3)) /= 0) then
    print *, "FAIL: want [0] got [", (f(3)), "]"
    stop 1
end if
contains
recursive integer function f(n) result(r)
integer::n
if (n==0) then
 r=0
else
 r=f(n-1)
end if
end function f
end program t
