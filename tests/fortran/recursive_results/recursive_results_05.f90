! vybe-test: fortran/recursive_results/recursive_results_05
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive logical function f() result(r)
r=.true.
end function f
