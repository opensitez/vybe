! vybe-test: fortran/procedure_results/procedure_results_02
! origin: languages/fortran/tests/fortran/test_procedure_results.rs
program t
if (nint((f()) * 100) /= 100) then
    print *, "FAIL: want [100] got [", nint((f()) * 100), "]"
    stop 1
end if
contains
real function f()
f=1.0
end function f
end program t
