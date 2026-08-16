! vybe-test: fortran/recursive_results/recursive_results_04
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if (trim(f()) /= "abc") then
    print *, "FAIL: want [abc] got [", trim(f()), "]"
    stop 1
end if
contains
recursive character(len=3) function f() result(r)
r='abc'
end function f
end program t
