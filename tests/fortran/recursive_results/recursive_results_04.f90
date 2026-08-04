! vybe-test: fortran/recursive_results/recursive_results_04
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive character(len=3) function f() result(r)
r='abc'
end function f
