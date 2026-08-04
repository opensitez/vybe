! vybe-test: fortran/full_programs/quadratic_discriminant
! origin: languages/fortran/tests/fortran/test_full_programs.rs
program t
integer :: vybe_check_i = 0
character(len=10) :: vybe_check_w(1) = [ "real roots" ]
real :: a, b, c, d
a = 1.0
b = -5.0
c = 6.0
d = b**2 - 4.0*a*c
if (d >= 0.0) then
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("real roots") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "real roots", "]"
    stop 1
end if
else
vybe_check_i = vybe_check_i + 1
if (vybe_check_i > 1) then
    print *, "FAIL: more than 1 line(s)"
    stop 1
end if
if (trim("complex roots") /= trim(vybe_check_w(vybe_check_i))) then
    print *, "FAIL at ", vybe_check_i, " got [", "complex roots", "]"
    stop 1
end if
end if
if (vybe_check_i /= 1) then
    print *, "FAIL: ", vybe_check_i, " line(s), wanted 1"
    stop 1
end if
end program t
