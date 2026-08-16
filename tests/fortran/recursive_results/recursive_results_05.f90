! vybe-test: fortran/recursive_results/recursive_results_05
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if (merge(1, 0, f()) /= 1) then
    print *, "FAIL: want [1] got [", merge(1, 0, f()), "]"
    stop 1
end if
contains
recursive logical function f() result(r)
r=.true.
end function f
end program t
