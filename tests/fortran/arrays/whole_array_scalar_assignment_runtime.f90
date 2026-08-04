! vybe-test: fortran/arrays/whole_array_scalar_assignment_runtime
! origin: languages/fortran/tests/fortran/test_arrays.rs

program test
    real :: a(3)
    a = 2.5
    if (abs((a(1)) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", a(1), "]"
    stop 1
end if
    if (abs((a(3)) - 2.5) > 1.0e-6) then
    print *, "FAIL: want [2.5] got [", a(3), "]"
    stop 1
end if
end program test
