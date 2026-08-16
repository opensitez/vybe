! vybe-test: fortran/recursive_results/recursive_results_06
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if ((f(3)) /= 3) then
    print *, "FAIL: want [3] got [", (f(3)), "]"
    stop 1
end if
contains
recursive integer function f(n) result(r)
integer::n
r=n
end function f
end program t
