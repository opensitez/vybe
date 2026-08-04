! vybe-test: fortran/trig_hyperbolic/atanh_roundtrip_from_tanh
! origin: languages/fortran/tests/fortran/test_trig_hyperbolic.rs
program t
real :: x
real :: y
x = 0.5
y = atanh(tanh(x))
if ((nint(y*1000)) /= 500) then
    print *, "FAIL: want [500] got [", nint(y*1000), "]"
    stop 1
end if
end program t
