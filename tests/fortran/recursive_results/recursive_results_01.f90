! vybe-test: fortran/recursive_results/recursive_results_01
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
recursive integer function f(n) result(r)
integer::n
if (n<=1) then
 r=1
else
 r=f(n-1)
end if
end function f
