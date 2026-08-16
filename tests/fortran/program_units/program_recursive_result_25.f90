! vybe-test: fortran/program_units/program_recursive_result_25
! origin: languages/fortran/tests/fortran/test_program_units.rs
program t
if (f(6) /= 720) then
    print *, "FAIL: want [720] got [", f(6), "]"
    stop 1
end if
contains
recursive integer function f(n) result(r)
integer :: n
if (n <= 1) then
 r = 1
else
 r = n * f(n-1)
end if
end function f
end program t
