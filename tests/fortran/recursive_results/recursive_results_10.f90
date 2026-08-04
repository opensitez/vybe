! vybe-test: fortran/recursive_results/recursive_results_10
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive integer function f(n) result(r)
integer::n
r=1
end function f
