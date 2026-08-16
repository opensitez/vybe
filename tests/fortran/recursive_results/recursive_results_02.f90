! vybe-test: fortran/recursive_results/recursive_results_02
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if (nint((f(1.5)) * 100) /= 150) then
    print *, "FAIL: want [150] got [", nint((f(1.5)) * 100), "]"
    stop 1
end if
contains
recursive real function f(x) result(r)
real::x
r=x
end function f
end program t
