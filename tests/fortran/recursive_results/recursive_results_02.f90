! vybe-test: fortran/recursive_results/recursive_results_02
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive real function f(x) result(r)
real::x
r=x
end function f
