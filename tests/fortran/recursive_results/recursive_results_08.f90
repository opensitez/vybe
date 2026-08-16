! vybe-test: fortran/recursive_results/recursive_results_08
! origin: languages/fortran/tests/fortran/test_recursive_results.rs
program t
if (nint((f()) * 100) /= 100) then
    print *, "FAIL: want [100] got [", nint((f()) * 100), "]"
    stop 1
end if
contains
recursive real function f() result(r)
r=1.0
end function f
end program t
