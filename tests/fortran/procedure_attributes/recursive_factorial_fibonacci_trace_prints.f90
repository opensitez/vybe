! vybe-test: fortran/procedure_attributes/recursive_factorial_fibonacci_trace_prints
! origin: languages/fortran/tests/fortran/test_procedure_attributes.rs
program t
integer :: vybe_check_i = 0
integer :: vybe_check_w(12) = [ 4, 3, 2, 1, 8, 5, 3, 2, 1, 1, 0, 8 ]
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 12) then
    print *, "FAIL: more than 12 line(s)"
    stop 1
end if
if ((fact_trace(4)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", fact_trace(4), "]"
    stop 1
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 12) then
    print *, "FAIL: more than 12 line(s)"
    stop 1
end if
if ((fib_trace(6)) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", fib_trace(6), "]"
    stop 1
end if
contains
recursive function fact_trace(n) result(r)
integer, intent(in) :: n
integer :: r
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 12) then
    print *, "FAIL: more than 12 line(s)"
    stop 1
end if
if ((n) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", n, "]"
    stop 1
end if
if (n <= 1) then
r = 1
else
r = n * fact_trace(n - 1)
end if
end function fact_trace
recursive function fib_trace(n) result(r)
integer, intent(in) :: n
integer :: r
if (n <= 1) then
r = n
else
r = fib_trace(n - 1) + fib_trace(n - 2)
end if
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 12) then
    print *, "FAIL: more than 12 line(s)"
    stop 1
end if
if ((r) /= vybe_check_w(vybe_check_i)) then
    print *, "FAIL at ", vybe_check_i, " got [", r, "]"
    stop 1
end if
end function fib_trace
if (vybe_check_i /= 12) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 12"
    stop 1
end if
end program t
