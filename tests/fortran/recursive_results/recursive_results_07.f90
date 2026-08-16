! vybe-test: fortran/recursive_results/recursive_results_07
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if ((f()) /= 1) then
    print *, "FAIL: want [1] got [", (f()), "]"
    stop 1
end if
contains
recursive integer function f() result(r)
r=1
end function f
end program t
