! vybe-test: fortran/procedure_results/procedure_results_07
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if (nint((f(1.5)) * 100) /= 150) then
    print *, "FAIL: want [150] got [", nint((f(1.5)) * 100), "]"
    stop 1
end if
contains
real function f(x)
real::x
f=x
end function f
end program t
